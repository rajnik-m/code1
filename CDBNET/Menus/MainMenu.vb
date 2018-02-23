Imports System.Drawing.Printing
Imports System.Linq

Public Class MainMenu

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)
  Private mvMenuStrip As MenuStrip
  Private mvToolStrip As ToolStrip
  Private mvParentForm As Form
  Private mvImageList16 As ImageList
  Private mvImageList32 As ImageList
  Private mvApplicationsMenu As ToolStripMenuItem
  Private mvImageProvider As ImageProvider
  Private mvToolbarTextPosition As TextImageRelation = TextImageRelation.ImageBeforeText

  Private Const MAX_DOCUMENT_COUNT As Integer = 50
  Private Const ACCESS_CONTROL_VERSION As Integer = 224
  'TRADER_APPLICATION_VERSION constant is in the CDBNetCL AppValues class

  Public Sub New(ByVal pParentForm As Form)
    pParentForm.SuspendLayout()
    InitMenus(pParentForm)
    pParentForm.ResumeLayout()
  End Sub

#Region "Private Methods"

  Private Sub InitMenus(ByVal pParentForm As Form)
    mvMenuStrip = New MenuStrip
    mvMenuStrip.Name = "MainMenuStrip"
    mvToolStrip = New ToolStrip
    mvToolStrip.AccessibleName = "Main Tool Bar"
    mvToolStrip.Name = "MainToolBar"
    mvToolStrip.AllowDrop = True

    mvMenuStrip.SuspendLayout()
    mvToolStrip.SuspendLayout()

    mvImageProvider = New ImageProvider
    mvImageProvider.Visible = False
    pParentForm.Controls.Add(mvImageProvider)

    pParentForm.Controls.Add(mvToolStrip)
    mvToolStrip.TabIndex = 1
    mvToolStrip.TabStop = True

    pParentForm.Controls.Add(mvMenuStrip)
    mvMenuStrip.TabIndex = 0
    pParentForm.MainMenuStrip = mvMenuStrip

    'pParentForm.Controls.SetChildIndex(mvMenuStrip, 0)
    'pParentForm.Controls.SetChildIndex(mvToolStrip, 1)

    AddHandler mvToolStrip.DragOver, AddressOf ToolbarDragOver
    AddHandler mvToolStrip.DragDrop, AddressOf ToolbarDragDrop
    AddHandler mvToolStrip.DoubleClick, AddressOf ToolbarDoubleClick

    mvParentForm = pParentForm

    mvImageList16 = mvImageProvider.NewImageList16
    mvImageList32 = mvImageProvider.NewImageList32
    BuildMenus()
    BuildToolbar()

    mvToolStrip.ResumeLayout()
    mvMenuStrip.ResumeLayout()

  End Sub

  Private Sub BuildMenus()
    mvApplicationsMenu = New ToolStripMenuItem
    Dim FileToolStripMenuItem As New ToolStripMenuItem
    Dim ViewToolStripMenuItem As New ToolStripMenuItem
    Dim QueryToolStripMenuItem As New ToolStripMenuItem
    Dim FindToolStripMenuItem As New ToolStripMenuItem
    Dim ToolsToolStripMenuItem As New ToolStripMenuItem
    Dim SystemToolStripMenuItem As New ToolStripMenuItem
    Dim AdminToolStripMenuItem As New ToolStripMenuItem
    Dim HelpToolStripMenuItem As New ToolStripMenuItem
    Dim WindowToolStripMenuItem As New ToolStripMenuItem
    mvMenuStrip.Items.AddRange(New ToolStripItem() {FileToolStripMenuItem,
                                   ViewToolStripMenuItem,
                                   QueryToolStripMenuItem,
                                   FindToolStripMenuItem,
                                   ToolsToolStripMenuItem,
                                   mvApplicationsMenu,
                                   SystemToolStripMenuItem,
                                   AdminToolStripMenuItem})
    mvMenuStrip.Items.Add(WindowToolStripMenuItem)
    If MDIForm IsNot Nothing Then
      mvMenuStrip.MdiWindowListItem = WindowToolStripMenuItem
    End If
    mvMenuStrip.Items.Add(HelpToolStripMenuItem)
    mvMenuStrip.TabIndex = 0

    FileToolStripMenuItem.Text = ControlText.MnuMFile
    ViewToolStripMenuItem.Text = ControlText.MnuMView
    QueryToolStripMenuItem.Text = ControlText.MnuMQuery
    FindToolStripMenuItem.Text = ControlText.MnuMFind
    ToolsToolStripMenuItem.Text = ControlText.MnuMTools
    mvApplicationsMenu.Text = ControlText.MnuMApplications
    SystemToolStripMenuItem.Text = ControlText.MnuMSystem
    AdminToolStripMenuItem.Text = ControlText.MnuMAdmin
    WindowToolStripMenuItem.Text = ControlText.MnuMWindow
    HelpToolStripMenuItem.Text = ControlText.MnuMHelp

    If MDIForm Is Nothing Then AddHandler WindowToolStripMenuItem.DropDownOpening, AddressOf WindowMenuOpeningHandler
    AddHandler QueryToolStripMenuItem.DropDownOpening, AddressOf QueryMenuOpeningHandler
    AddHandler SystemToolStripMenuItem.DropDownOpening, AddressOf SystemMenuOpeningHandler

    With mvMenuItems
      .Add(CommandIndexes.cbiLogin.ToString, New MenuToolbarCommand("Login", ControlText.MnuMLogin, CommandIndexes.cbiLogin, , mvImageList16.Images(CommandIndexes.cbiLogin), ControlText.MnuMLoginTT))
      .Add(CommandIndexes.cbiNewContact.ToString, New MenuToolbarCommand("NewContact", ControlText.MnuMNewContact, CommandIndexes.cbiNewContact, "SCFLNC", mvImageList16.Images(CommandIndexes.cbiNewContact), "New Contact"))
      .Add(CommandIndexes.cbiNewContact2.ToString, New MenuToolbarCommand("NewContact2", "&Contact2...", CommandIndexes.cbiNewContact2, "SCFLC2", mvImageList16.Images(CommandIndexes.cbiNewContact2), "New Contact2"))
      .Add(CommandIndexes.cbiNewContact3.ToString, New MenuToolbarCommand("NewContact3", "&Contact3...", CommandIndexes.cbiNewContact3, "SCFLC3", mvImageList16.Images(CommandIndexes.cbiNewContact3), "New Contact3"))
      .Add(CommandIndexes.cbiNewContact4.ToString, New MenuToolbarCommand("NewContact4", "&Contact4...", CommandIndexes.cbiNewContact4, "SCFLC4", mvImageList16.Images(CommandIndexes.cbiNewContact4), "New Contact4"))
      .Add(CommandIndexes.cbiNewContact5.ToString, New MenuToolbarCommand("NewContact5", "&Contact5...", CommandIndexes.cbiNewContact5, "SCFLC5", mvImageList16.Images(CommandIndexes.cbiNewContact5), "New Contact5"))
      .Add(CommandIndexes.cbiNewOrganisation.ToString, New MenuToolbarCommand("NewOrganisation", ControlText.MnuMNewOrganisation, CommandIndexes.cbiNewOrganisation, "SCFLNO", mvImageList16.Images(CommandIndexes.cbiNewOrganisation), "New Organisation"))
      .Add(CommandIndexes.cbiNewOrganisation2.ToString, New MenuToolbarCommand("NewOrganisation2", "&Organisation2...", CommandIndexes.cbiNewOrganisation2, "SCFLO2", mvImageList16.Images(CommandIndexes.cbiNewOrganisation2), "New Organisation2"))
      .Add(CommandIndexes.cbiNewOrganisation3.ToString, New MenuToolbarCommand("NewOrganisation3", "&Organisation3...", CommandIndexes.cbiNewOrganisation3, "SCFLO3", mvImageList16.Images(CommandIndexes.cbiNewOrganisation3), "New Organisation3"))
      .Add(CommandIndexes.cbiNewOrganisation4.ToString, New MenuToolbarCommand("NewOrganisation4", "&Organisation4...", CommandIndexes.cbiNewOrganisation4, "SCFLO4", mvImageList16.Images(CommandIndexes.cbiNewOrganisation4), "New Organisation4"))
      .Add(CommandIndexes.cbiNewOrganisation5.ToString, New MenuToolbarCommand("NewOrganisation5", "&Organisation5...", CommandIndexes.cbiNewOrganisation5, "SCFLO5", mvImageList16.Images(CommandIndexes.cbiNewOrganisation5), "New Organisation5"))
      .Add(CommandIndexes.cbiNewDocument.ToString, New MenuToolbarCommand("NewDocument", ControlText.MnuMNewDocument, CommandIndexes.cbiNewDocument, "SCFLND", mvImageList16.Images(CommandIndexes.cbiNewDocument), ControlText.MnuMNewDocumentTT))
      .Add(CommandIndexes.cbiNewAction.ToString, New MenuToolbarCommand("NewAction", ControlText.MnuMNewAction, CommandIndexes.cbiNewAction, "SCFLNA", mvImageList16.Images(CommandIndexes.cbiNewAction), ControlText.MnuMNewActionTT))
      .Add(CommandIndexes.cbiNewActionTemplate.ToString, New MenuToolbarCommand("NewActionTemplate", ControlText.MnuMNewActionTemplate, CommandIndexes.cbiNewActionTemplate, "SCFLNB", mvImageList16.Images(CommandIndexes.cbiNewAction), ControlText.MnuMNewActionTemplateTT))
      .Add(CommandIndexes.cbiNewTelephoneCall.ToString, New MenuToolbarCommand("NewTelephoneCall", ControlText.MnuMNewTelephoneCall, CommandIndexes.cbiNewTelephoneCall, "SCFLNT", mvImageList16.Images(CommandIndexes.cbiNewTelephoneCall), ControlText.MnuMNewTelephoneCallTT))
      .Add(CommandIndexes.cbiNewSelectionSet.ToString, New MenuToolbarCommand("NewSelectionSet", ControlText.MnuMNewSelectionSet, CommandIndexes.cbiNewSelectionSet, "SCFLNS", mvImageList16.Images(CommandIndexes.cbiNewSelectionSet), ControlText.MnuMNewSelectionSetTT))
      .Add(CommandIndexes.cbiNewEvent.ToString, New MenuToolbarCommand("NewEvent", ControlText.MnuMNewEvent, CommandIndexes.cbiNewEvent, "SCFLEV", mvImageList16.Images(CommandIndexes.cbiNewEvent), "New Event"))
      .Add(CommandIndexes.cbiNewEvent2.ToString, New MenuToolbarCommand("NewEvent2", "&Event2...", CommandIndexes.cbiNewEvent2, "SCFLE2", mvImageList16.Images(CommandIndexes.cbiNewEvent2), "New Event2"))
      .Add(CommandIndexes.cbiNewEvent3.ToString, New MenuToolbarCommand("NewEvent3", "&Event3...", CommandIndexes.cbiNewEvent3, "SCFLE3", mvImageList16.Images(CommandIndexes.cbiNewEvent3), "New Event3"))
      .Add(CommandIndexes.cbiNewEvent4.ToString, New MenuToolbarCommand("NewEvent4", "&Event4...", CommandIndexes.cbiNewEvent4, "SCFLE4", mvImageList16.Images(CommandIndexes.cbiNewEvent4), "New Event4"))
      .Add(CommandIndexes.cbiNewEvent5.ToString, New MenuToolbarCommand("NewEvent5", "&Event5...", CommandIndexes.cbiNewEvent5, "SCFLE5", mvImageList16.Images(CommandIndexes.cbiNewEvent5), "New Event5"))
      .Add(CommandIndexes.cbiPreferences.ToString, New MenuToolbarCommand("Preferences", ControlText.MnuMPreferences, CommandIndexes.cbiPreferences, "SCLMPR", mvImageList16.Images(CommandIndexes.cbiPreferences), ControlText.MnuMPreferencesTT))
      .Add(CommandIndexes.cbiPageSetup.ToString, New MenuToolbarCommand("PageSetup", ControlText.MnuMPageSetup, CommandIndexes.cbiPageSetup, "SCLMPS", mvImageList16.Images(CommandIndexes.cbiPageSetup), ControlText.MnuMPageSetupTT))
      .Add(CommandIndexes.cbiExit.ToString, New MenuToolbarCommand("Exit", ControlText.MnuMExit, CommandIndexes.cbiExit, , mvImageList16.Images(CommandIndexes.cbiExit), ControlText.MnuMExit))
      .Add(CommandIndexes.cbiLogWEBServices.ToString, New MenuToolbarCommand("LogWEBServices", ControlText.MnuMLogWebServices, CommandIndexes.cbiLogWEBServices, "SCLMLW", , ControlText.MnuMLogWebServiceCallsTT))

      .Add(CommandIndexes.cbiNextRecord.ToString, New MenuToolbarCommand("NextRecord", "Next Record", CommandIndexes.cbiNextRecord, "SCVMTB", mvImageList16.Images(CommandIndexes.cbiNextRecord), "Next Record"))
      .Add(CommandIndexes.cbiPreviousRecord.ToString, New MenuToolbarCommand("PreviousRecord", "Previous Record", CommandIndexes.cbiPreviousRecord, "SCVMTB", mvImageList16.Images(CommandIndexes.cbiPreviousRecord), "Previous Record"))
      .Add(CommandIndexes.cbiToolbar.ToString, New MenuToolbarCommand("Toolbar", ControlText.MnuMToolbar, CommandIndexes.cbiToolbar, "SCVMTB", , ControlText.MnuMToolbarTT))
      .Add(CommandIndexes.cbiNavigationPanel.ToString, New MenuToolbarCommand("NavigationPanel", ControlText.MnuMNavPanel, CommandIndexes.cbiNavigationPanel, "SCVMNP", , ControlText.MnuMNavPanelTT))
      .Add(CommandIndexes.cbiStatusBar.ToString, New MenuToolbarCommand("StatusBar", ControlText.MnuMStatusBar, CommandIndexes.cbiStatusBar, "SCVMSB", , ControlText.MnuMStatusBarTT))
      .Add(CommandIndexes.cbiHeaderPanel.ToString, New MenuToolbarCommand("HeaderPanel", ControlText.MnuMHeaderPanel, CommandIndexes.cbiHeaderPanel, "SCVMHP", , ControlText.MnuMHeaderPanelTT))
      .Add(CommandIndexes.cbiSelectionPanel.ToString, New MenuToolbarCommand("SelectionPanel", ControlText.MnuMSelectionPanel, CommandIndexes.cbiSelectionPanel, "SCVMSP", , ControlText.MnuMSelectionPanelTT))
      .Add(CommandIndexes.cbiDashboard.ToString, New MenuToolbarCommand("Dashboard", "Dashboard", CommandIndexes.cbiDashboard, "SCVMDB", mvImageList16.Images(CommandIndexes.cbiDashboard), "Dashboard"))

      .Add(CommandIndexes.cbiMyDetails.ToString, New MenuToolbarCommand("MyDetails", ControlText.MnuMMyDetails, CommandIndexes.cbiMyDetails, "SCVMMT", mvImageList16.Images(CommandIndexes.cbiMyDetails), ControlText.MnuMMyDetailsTT))
      .Add(CommandIndexes.cbiMyOrganisation.ToString, New MenuToolbarCommand("MyOrganisation", ControlText.MnuMMyOrganisation, CommandIndexes.cbiMyOrganisation, "SCVMMO", mvImageList16.Images(CommandIndexes.cbiMyOrganisation), ControlText.MnuMMyOrganisationTT))
      .Add(CommandIndexes.cbiMyActions.ToString, New MenuToolbarCommand("MyActions", ControlText.MnuMMyActions, CommandIndexes.cbiMyActions, "SCVMMA", mvImageList16.Images(CommandIndexes.cbiMyActions), ControlText.MnuMMyActionsTT))
      .Add(CommandIndexes.cbiMyDocuments.ToString, New MenuToolbarCommand("MyDocuments", ControlText.MnuMMyDocuments, CommandIndexes.cbiMyDocuments, "SCVMMD", mvImageList16.Images(CommandIndexes.cbiMyDocuments), ControlText.MnuMMyDocumentsTT))
      .Add(CommandIndexes.cbiMyInBox.ToString, New MenuToolbarCommand("MyInbox", ControlText.MnuMMyInbox, CommandIndexes.cbiMyInBox, "SCVMMI", mvImageList16.Images(CommandIndexes.cbiMyInBox), ControlText.MnuMMyInboxTT))
      .Add(CommandIndexes.cbiMyJournal.ToString, New MenuToolbarCommand("MyJournal", ControlText.MnuMMyJournal, CommandIndexes.cbiMyJournal, "SCVMMJ", mvImageList16.Images(CommandIndexes.cbiMyJournal), ControlText.MnuMMyJournalTT))
      .Add(CommandIndexes.cbiRefresh.ToString, New MenuToolbarCommand("Refresh", ControlText.MnuMRefresh, CommandIndexes.cbiRefresh, "SCVMRE", mvImageList16.Images(CommandIndexes.cbiRefresh), ControlText.MnuMRefreshTT))

      .Add(CommandIndexes.cbiSearchData.ToString, New MenuToolbarCommand("SearchData", ControlText.MnuMSearchData, CommandIndexes.cbiSearchData, "SCFMSE", mvImageList16.Images(CommandIndexes.cbiSearchData), ControlText.MnuMSearchDataTT))
      .Add(CommandIndexes.cbiContactFinder.ToString, New MenuToolbarCommand("FindContact", ControlText.MnuMFindContacts, CommandIndexes.cbiContactFinder, "SCFMPF", mvImageList16.Images(CommandIndexes.cbiContactFinder), "Find Contacts"))
      .Add(CommandIndexes.cbiContactFinder2.ToString, New MenuToolbarCommand("FindContact2", "&Contacts2...", CommandIndexes.cbiContactFinder2, "SCFMP2", mvImageList16.Images(CommandIndexes.cbiContactFinder2), "Find Contacts2"))
      .Add(CommandIndexes.cbiContactFinder3.ToString, New MenuToolbarCommand("FindContact3", "&Contacts3...", CommandIndexes.cbiContactFinder3, "SCFMP3", mvImageList16.Images(CommandIndexes.cbiContactFinder3), "Find Contacts3"))
      .Add(CommandIndexes.cbiContactFinder4.ToString, New MenuToolbarCommand("FindContact4", "&Contacts4...", CommandIndexes.cbiContactFinder4, "SCFMP4", mvImageList16.Images(CommandIndexes.cbiContactFinder4), "Find Contacts4"))
      .Add(CommandIndexes.cbiContactFinder5.ToString, New MenuToolbarCommand("FindContact5", "&Contacts5...", CommandIndexes.cbiContactFinder5, "SCFMP5", mvImageList16.Images(CommandIndexes.cbiContactFinder5), "Find Contacts5"))
      .Add(CommandIndexes.cbiOrganisationFinder.ToString, New MenuToolbarCommand("FindOrganisation", ControlText.MnuMFindOrganisations, CommandIndexes.cbiOrganisationFinder, "SCFMOF", mvImageList16.Images(CommandIndexes.cbiOrganisationFinder), "Find Organisations"))
      .Add(CommandIndexes.cbiOrganisationFinder2.ToString, New MenuToolbarCommand("FindOrganisation2", "&Organisations2...", CommandIndexes.cbiOrganisationFinder2, "SCFMO2", mvImageList16.Images(CommandIndexes.cbiOrganisationFinder2), "Find Organisations2"))
      .Add(CommandIndexes.cbiOrganisationFinder3.ToString, New MenuToolbarCommand("FindOrganisation3", "&Organisations3...", CommandIndexes.cbiOrganisationFinder3, "SCFMO3", mvImageList16.Images(CommandIndexes.cbiOrganisationFinder3), "Find Organisations3"))
      .Add(CommandIndexes.cbiOrganisationFinder4.ToString, New MenuToolbarCommand("FindOrganisation4", "&Organisations4...", CommandIndexes.cbiOrganisationFinder4, "SCFMO4", mvImageList16.Images(CommandIndexes.cbiOrganisationFinder4), "Find Organisations4"))
      .Add(CommandIndexes.cbiOrganisationFinder5.ToString, New MenuToolbarCommand("FindOrganisation5", "&Organisations5...", CommandIndexes.cbiOrganisationFinder5, "SCFMO5", mvImageList16.Images(CommandIndexes.cbiOrganisationFinder5), "Find Organisations5"))
      .Add(CommandIndexes.cbiDocumentFinder.ToString, New MenuToolbarCommand("FindDocument", ControlText.MnuMFindDocuments, CommandIndexes.cbiDocumentFinder, "SCFMDF", mvImageList16.Images(CommandIndexes.cbiDocumentFinder), ControlText.MnuMFindDocumentsTT))
      .Add(CommandIndexes.cbiMeeting.ToString, New MenuToolbarCommand("FindMeeting", ControlText.MnuMFindMeeting, CommandIndexes.cbiMeeting, "SCFMMF", mvImageList16.Images(CommandIndexes.cbiMeeting), ControlText.MnuMFindMeeting))
      .Add(CommandIndexes.cbiActionFinder.ToString, New MenuToolbarCommand("FindAction", ControlText.MnuMFindActions, CommandIndexes.cbiActionFinder, "SCFMAF", mvImageList16.Images(CommandIndexes.cbiActionFinder), ControlText.MnuMFindActionsTT))
      .Add(CommandIndexes.cbiSelectionSetFinder.ToString, New MenuToolbarCommand("FindSelectionSets", ControlText.MnuMFindSelectionSets, CommandIndexes.cbiSelectionSetFinder, "SCFMSS", mvImageList16.Images(CommandIndexes.cbiSelectionSetFinder), ControlText.MnuMFindSelectionSetsTT))
      .Add(CommandIndexes.cbiEventFinder.ToString, New MenuToolbarCommand("FindEvent", ControlText.MnuMFindEvents, CommandIndexes.cbiEventFinder, "SCFMEF", mvImageList16.Images(CommandIndexes.cbiEventFinder), "Find Events"))
      .Add(CommandIndexes.cbiEventFinder2.ToString, New MenuToolbarCommand("FindEvent2", "&Events2...", CommandIndexes.cbiEventFinder2, "SCFME2", mvImageList16.Images(CommandIndexes.cbiEventFinder2), "Find Events2"))
      .Add(CommandIndexes.cbiEventFinder3.ToString, New MenuToolbarCommand("FindEvent3", "&Events3...", CommandIndexes.cbiEventFinder3, "SCFME3", mvImageList16.Images(CommandIndexes.cbiEventFinder3), "Find Events3"))
      .Add(CommandIndexes.cbiEventFinder4.ToString, New MenuToolbarCommand("FindEvent4", "&Events4...", CommandIndexes.cbiEventFinder4, "SCFME4", mvImageList16.Images(CommandIndexes.cbiEventFinder4), "Find Events4"))
      .Add(CommandIndexes.cbiEventFinder5.ToString, New MenuToolbarCommand("FindEvent5", "&Events5...", CommandIndexes.cbiEventFinder5, "SCFME5", mvImageList16.Images(CommandIndexes.cbiEventFinder5), "Find Events5"))
      .Add(CommandIndexes.cbiMemberFinder.ToString, New MenuToolbarCommand("FindMember", ControlText.MnuMFindMembers, CommandIndexes.cbiMemberFinder, "SCFMBF", mvImageList16.Images(CommandIndexes.cbiMemberFinder), ControlText.MnuMFindMembersTT))
      .Add(CommandIndexes.cbiPayPlanFinder.ToString, New MenuToolbarCommand("FindPaymentPlan", ControlText.MnuMFindPaymentPlans, CommandIndexes.cbiPayPlanFinder, "SCFMPP", mvImageList16.Images(CommandIndexes.cbiPayPlanFinder), ControlText.MnuMFindPaymentPlansTT))
      .Add(CommandIndexes.cbiCovenantFinder.ToString, New MenuToolbarCommand("FindCovenant", ControlText.MnuMFindCovenants, CommandIndexes.cbiCovenantFinder, "SCFMCF", mvImageList16.Images(CommandIndexes.cbiCovenantFinder), ControlText.MnuMFindCovenantsTT))
      .Add(CommandIndexes.cbiXactionFinder.ToString, New MenuToolbarCommand("FindTransaction", ControlText.MnuMFindTransactions, CommandIndexes.cbiXactionFinder, "SCFMTF", mvImageList16.Images(CommandIndexes.cbiXactionFinder), ControlText.MnuMFindTransactionsTT))
      .Add(CommandIndexes.cbiStandingOrderFinder.ToString, New MenuToolbarCommand("FindStandingOrder", ControlText.MnuMFindStandingOrders, CommandIndexes.cbiStandingOrderFinder, "SCFMSO", mvImageList16.Images(CommandIndexes.cbiStandingOrderFinder), ControlText.MnuMFindStandingOrdersTT))
      .Add(CommandIndexes.cbiDirectDebitFinder.ToString, New MenuToolbarCommand("FindDirectDebit", ControlText.MnuMFindDirectDebits, CommandIndexes.cbiDirectDebitFinder, "SCFMDD", mvImageList16.Images(CommandIndexes.cbiDirectDebitFinder), ControlText.MnuMFindDirectDebitsTT))
      .Add(CommandIndexes.cbiCCCAFinder.ToString, New MenuToolbarCommand("FindCCCA", ControlText.MnuMFindCCCAs, CommandIndexes.cbiCCCAFinder, "SCFMCC", mvImageList16.Images(CommandIndexes.cbiCCCAFinder), ControlText.MnuMFindCCCAsTT))
      .Add(CommandIndexes.cbiGADFinder.ToString, New MenuToolbarCommand("FindGAD", ControlText.MnuMFindGiftAid, CommandIndexes.cbiGADFinder, "SCFMGA", mvImageList16.Images(CommandIndexes.cbiGADFinder), ControlText.MnuMFindGiftAidTT))
      .Add(CommandIndexes.cbiInvoiceFinder.ToString, New MenuToolbarCommand("FindInvoiceCreditNote", ControlText.MnuMFindInvoices, CommandIndexes.cbiInvoiceFinder, "SCFMIF", mvImageList16.Images(CommandIndexes.cbiInvoiceFinder), ControlText.MnuMFindInvoicesTT))
      .Add(CommandIndexes.cbiLegacyFinder.ToString, New MenuToolbarCommand("FindLegacy", ControlText.MnuMFindLegacies, CommandIndexes.cbiLegacyFinder, "SCFMLG", mvImageList16.Images(CommandIndexes.cbiLegacyFinder), ControlText.MnuMFindLegaciesTT))
      .Add(CommandIndexes.cbiGAYEFinder.ToString, New MenuToolbarCommand("FindPreTaxPG", ControlText.MnuMFindPreTaxPG, CommandIndexes.cbiGAYEFinder, "SCFMGY", mvImageList16.Images(CommandIndexes.cbiGAYEFinder), ControlText.MnuMFindPreTaxPGTT))
      .Add(CommandIndexes.cbiPostTaxPGFinder.ToString, New MenuToolbarCommand("FindPostTaxPG", ControlText.MnuMFindPostTaxPG, CommandIndexes.cbiPostTaxPGFinder, "SCFMPG", mvImageList16.Images(CommandIndexes.cbiGAYEFinder), ControlText.MnuMFindPostTaxPGTT))
      .Add(CommandIndexes.cbiPurchaseOrderFinder.ToString, New MenuToolbarCommand("FindPurchaseOrder", ControlText.MnuMFindPurchaseOrders, CommandIndexes.cbiPurchaseOrderFinder, "SCFMPO", mvImageList16.Images(CommandIndexes.cbiPurchaseOrderFinder), ControlText.MnuMFindPurchaseOrdersTT))
      .Add(CommandIndexes.cbiProductFinder.ToString, New MenuToolbarCommand("FindProduct", ControlText.MnuMFindProducts, CommandIndexes.cbiProductFinder, "SCFMPC", mvImageList16.Images(CommandIndexes.cbiProductFinder), ControlText.MnuMFindProductsTT))
      .Add(CommandIndexes.cbiCampaignFinder.ToString, New MenuToolbarCommand("FindCampaign", ControlText.MnuMFindCampaigns, CommandIndexes.cbiCampaignFinder, "SCFMCA", mvImageList16.Images(CommandIndexes.cbiCampaignFinder), ControlText.MnuMFindCampaignsTT))
      .Add(CommandIndexes.cbiStandardDocument.ToString, New MenuToolbarCommand("FindStandardDocument", ControlText.MnuMFindStandardDocuments, CommandIndexes.cbiStandardDocument, "SCFMSD", mvImageList16.Images(CommandIndexes.cbiStandardDocument), ControlText.MnuMFindStandardDocumentsTT))
      .Add(CommandIndexes.cbiFundraisingPaymentScheduleFinder.ToString, New MenuToolbarCommand("FindFundraisingPayments", ControlText.MnuMFindFundraisingPayments, CommandIndexes.cbiFundraisingPaymentScheduleFinder, "SCFMFP", , ControlText.MnuMFindFundraisingPayments))
      .Add(CommandIndexes.cbiServiceProductFinder.ToString, New MenuToolbarCommand("FindServiceProduct", ControlText.MnuMFindServiceProducts, CommandIndexes.cbiServiceProductFinder, "SCFMSP", , ControlText.MnuMFindServiceProductsTT))
      .Add(CommandIndexes.cbiFundraisingRequestFinder.ToString, New MenuToolbarCommand("FindFundraisingRequest", ControlText.MnuMFindFundraisingRequests, CommandIndexes.cbiFundraisingRequestFinder, "SCFMFR", , ControlText.MnuMFindFundraisingRequests))
      .Add(CommandIndexes.cbiSeparator.ToString, New MenuToolbarCommand("Separator", CommandIndexes.cbiSeparator))

      .Add(CommandIndexes.cbiTableMaintenance.ToString, New MenuToolbarCommand("TableMaintenance", ControlText.MnuMTableMaintenance, CommandIndexes.cbiTableMaintenance, "SCTMTM", mvImageList16.Images(CommandIndexes.cbiTableMaintenance), ControlText.MnuMTableMaintenanceTT))
      .Add(CommandIndexes.cbiJobSchedule.ToString, New MenuToolbarCommand("JobSchedule", ControlText.MnuMJobSchedule, CommandIndexes.cbiJobSchedule, "SCTMJS", mvImageList16.Images(CommandIndexes.cbiJobSchedule), ControlText.MnuMJobScheduleTT))
      .Add(CommandIndexes.cbiListManager.ToString, New MenuToolbarCommand("ListManager", ControlText.MnuMListManager, CommandIndexes.cbiListManager, "SCTMLM", mvImageList16.Images(CommandIndexes.cbiListManager), ControlText.MnuMListManagerTT))
      .Add(CommandIndexes.cbiDocumentDistributor.ToString, New MenuToolbarCommand("Document Distributor", ControlText.MnuDocumentDistributor, CommandIndexes.cbiDocumentDistributor, "SCTMDD", mvImageList16.Images(CommandIndexes.cbiDocumentDistributor), ControlText.MnuDocumentDistributorTT))
      .Add(CommandIndexes.cbiSendEmail.ToString, New MenuToolbarCommand("SendEMail", ControlText.MnuMSendEmail, CommandIndexes.cbiSendEmail, "SCTMSM", mvImageList16.Images(CommandIndexes.cbiSendEmail), ControlText.MnuMSendEmailTT))
      .Add(CommandIndexes.cbiRunTest.ToString, New MenuToolbarCommand("RunTest", "Run Test", CommandIndexes.cbiRunTest, , , "Run Test"))
      .Add(CommandIndexes.cbiClearCache.ToString, New MenuToolbarCommand("Clear Cache", ControlText.MnuMClearCache, CommandIndexes.cbiClearCache, , , ControlText.MnuMClearCacheTT))
      .Add(CommandIndexes.cbiAllowCaching.ToString, New MenuToolbarCommand("Allow Caching", "Allow Caching", CommandIndexes.cbiAllowCaching, , , ))
      .Add(CommandIndexes.cbiExplore.ToString, New MenuToolbarCommand("Explore", ControlText.MnuMExplore, CommandIndexes.cbiExplore, "SCTMEX", , "Explore"))
      .Add(CommandIndexes.cbiCustomise.ToString, New MenuToolbarCommand("Customise", ControlText.MnuMCustomise, CommandIndexes.cbiCustomise, "SCTMCU", mvImageList16.Images(CommandIndexes.cbiCustomise), ControlText.MnuMCustomiseTT))
      .Add(CommandIndexes.cbiRunMailing.ToString, New MenuToolbarCommand("RunMailing", ControlText.MnuMRunMailing, CommandIndexes.cbiRunMailing, "SCTMRM", mvImageList16.Images(CommandIndexes.cbiRunMailing), ControlText.MnuMRunMailingTT))
      .Add(CommandIndexes.cbiRunReport.ToString, New MenuToolbarCommand("RunReport", ControlText.MnuMRunReport, CommandIndexes.cbiRunReport, "SCTMRR", mvImageList16.Images(CommandIndexes.cbiRunReport), ControlText.MnuMRunReportTT))
      .Add(CommandIndexes.cbiCloseOpenBatch.ToString, New MenuToolbarCommand("CloseOpenBatch", ControlText.MnuMCloseBatch, CommandIndexes.cbiCloseOpenBatch, "SCTMCO", mvImageList16.Images(CommandIndexes.cbiCloseOpenBatch), ControlText.MnuMCloseBatchTT))
      .Add(CommandIndexes.cbiCopyEventPricingMatrix.ToString, New MenuToolbarCommand("CopyEventPriceMatrix", ControlText.MnuMCopyEventPriceMatrix, CommandIndexes.cbiCopyEventPricingMatrix, "SCTMCM", , ControlText.MnuMCopyEventPriceMatrixTT))
      .Add(CommandIndexes.cbiPostcodeProximity.ToString, New MenuToolbarCommand("PostcodeProximity", ControlText.MnuAdminPostcodeProximity, CommandIndexes.cbiPostcodeProximity, "SCTMPP", , ControlText.MnuAdminPostcodeProximityTT, False))

      .Add(CommandIndexes.cbiCascade.ToString, New MenuToolbarCommand("Cascade", ControlText.MnuMCascade, CommandIndexes.cbiCascade, , mvImageList16.Images(CommandIndexes.cbiCascade), ControlText.MnuMCascadeTT))
      .Add(CommandIndexes.cbiTileHorizontally.ToString, New MenuToolbarCommand("TileHorizontally", ControlText.MnuMTileH, CommandIndexes.cbiTileHorizontally, , mvImageList16.Images(CommandIndexes.cbiTileHorizontally), ControlText.MnuMTileHTT))
      .Add(CommandIndexes.cbiTileVertically.ToString, New MenuToolbarCommand("TileVertically", ControlText.MnuMTileV, CommandIndexes.cbiTileVertically, , mvImageList16.Images(CommandIndexes.cbiTileVertically), ControlText.MnuMTileVTT))
      .Add(CommandIndexes.cbiArrangeIcons.ToString, New MenuToolbarCommand("ArrangeAll", ControlText.MnuMArrangeAll, CommandIndexes.cbiArrangeIcons, , mvImageList16.Images(CommandIndexes.cbiArrangeIcons), ControlText.MnuMArrangeTT))
      .Add(CommandIndexes.cbiCloseAll.ToString, New MenuToolbarCommand("CloseAll", ControlText.MnuMCloseAll, CommandIndexes.cbiCloseAll, , mvImageList16.Images(CommandIndexes.cbiCloseAll), ControlText.MnuMCloseAllTT))

      .Add(CommandIndexes.cbiHelp.ToString, New MenuToolbarCommand("Help", ControlText.MnuMHelpContents, CommandIndexes.cbiHelp, , , ControlText.MnuMHelpTT))
      .Add(CommandIndexes.cbiReleaseNotes.ToString, New MenuToolbarCommand("Release Notes", ControlText.MnuMHelpReleaseNotes, CommandIndexes.cbiReleaseNotes,,, ControlText.MnuMHelpReleaseNotesTT))
      .Add(CommandIndexes.cbiKnowledgebase.ToString, New MenuToolbarCommand("Knowledgebase", ControlText.MnuMKnowledgebase, CommandIndexes.cbiKnowledgebase, , , ControlText.MnuMKnowledgebaseTT))
      .Add(CommandIndexes.cbiSupportForum.ToString, New MenuToolbarCommand("Support Forum", ControlText.MnuMSupportForum, CommandIndexes.cbiSupportForum, , , ControlText.MnuMSupportForumTT))
      .Add(CommandIndexes.cbiAbout.ToString, New MenuToolbarCommand("About", GetInformationMessage(ControlText.MnuMAbout), CommandIndexes.cbiAbout, , , GetInformationMessage(ControlText.MnuMAboutTT)))

      '-------------------------------------------------------------
      'SYSTEM MENU
      '-------------------------------------------------------------

      .Add(CommandIndexes.mnuFinBanksLoad.ToString, New MenuToolbarCommand("LoadBankData", ControlText.MnuSLoadBankData, CommandIndexes.mnuFinBanksLoad, "SCBALD", , ControlText.MnuSLoadBankDataTT, True))

      .Add(CommandIndexes.mnuFinBViewBatchDetail.ToString, New MenuToolbarCommand("ViewBatchDetails", ControlText.MnuSViewBatchDetails, CommandIndexes.mnuFinBViewBatchDetail, "SCSMVB", , ControlText.MnuSViewBatchDetailsTT, True))
      .Add(CommandIndexes.mnuFinBProcessBatches.ToString, New MenuToolbarCommand("ProcessBatches", ControlText.MnuSProcessBatches, CommandIndexes.mnuFinBProcessBatches, "SCSMPB", , ControlText.MnuSProcessBatchesTT, True))
      .Add(CommandIndexes.mnuFinBOutstandingBatchesReport.ToString, New MenuToolbarCommand("OutstandingBatchesReport", ControlText.MnuSOutstandingBatchesReport, CommandIndexes.mnuFinBOutstandingBatchesReport, "SCSMOS", , ControlText.MnuSOutstandingBatchesReportTT, True))
      .Add(CommandIndexes.mnuFinBChequeList.ToString, New MenuToolbarCommand("PrintChequeList", ControlText.MnuSPrintChequeList, CommandIndexes.mnuFinBChequeList, "SCSMPC", , ControlText.MnuSPrintChequeListTT, True))
      .Add(CommandIndexes.mnuFinBPayingInSlips.ToString, New MenuToolbarCommand("PrintPayingInSlip", ControlText.MnuSPrintPayingInSlip, CommandIndexes.mnuFinBPayingInSlips, "SCSMPS", , ControlText.MnuSPrintPayingInSlipTT, True))
      .Add(CommandIndexes.mnuFinBCashBookBatch.ToString, New MenuToolbarCommand("RedoCashBookBatch", ControlText.MnuSRedoCashBookBatch, CommandIndexes.mnuFinBCashBookBatch, "SCSMCB", , ControlText.MnuSRedoCashBookBatchTT, True))
      .Add(CommandIndexes.mnuFinBUpdateBatch.ToString, New MenuToolbarCommand("PostBatch", ControlText.MnuSPostBatch, CommandIndexes.mnuFinBUpdateBatch, "SCSMPO", , ControlText.MnuSPostBatchTT, True))
      .Add(CommandIndexes.mnuFinBCreateJournal.ToString, New MenuToolbarCommand("CreateJournalFiles", ControlText.MnuSCreateJournalFiles, CommandIndexes.mnuFinBCreateJournal, "SCSMCJ", , ControlText.MnuSCreateJournalFilesTT, True))
      .Add(CommandIndexes.mnuFinBSummaryReport.ToString, New MenuToolbarCommand("SummaryReport", ControlText.MnuSSummaryReport, CommandIndexes.mnuFinBSummaryReport, "SCSMSR", , ControlText.MnuSSummaryReportTT, True))
      .Add(CommandIndexes.mnuFinBDetailReport.ToString, New MenuToolbarCommand("DetailReport", ControlText.MnuSDetailReport, CommandIndexes.mnuFinBDetailReport, "SCSMDT", , ControlText.MnuSDetailReportTT, True))
      .Add(CommandIndexes.mnuFinBPurgeOldBatches.ToString, New MenuToolbarCommand("PurgeOldBatches", ControlText.MnuSPurgeOldBatches, CommandIndexes.mnuFinBPurgeOldBatches, "SCSMPU", , ControlText.MnuSPurgeOldBatchesTT, True))
      'Removed close open batch as it is on the tools menu
      '.Add(CommandIndexes.mnuFinBCloseOpenBatch.ToString, New MenuToolbarCommand("CloseOpenBatch", ControlText.MnuSCloseOpenBatch, CommandIndexes.mnuFinBCloseOpenBatch, "SCSMOB", , ControlText.MnuSCloseOpenBatchTT, True))
      .Add(CommandIndexes.mnuFinBPurgePrizeDrawBatches.ToString, New MenuToolbarCommand("PurgePrizeDrawBatches", ControlText.MnuSPurgePrizeDrawBatches, CommandIndexes.mnuFinBPurgePrizeDrawBatches, "SCSMPZ", , ControlText.MnuSPurgePrizeDrawBatchesTT, True))

      .Add(CommandIndexes.mnuFinCAFExpectedPayments.ToString, New MenuToolbarCommand("SO/CCCAExpectedPaymentsReport", ControlText.MnuSSOCCCAExpectedPaymentsReport, CommandIndexes.mnuFinCAFExpectedPayments, "SCFCEP", , ControlText.MnuSSOCCCAExpectedPaymentsReportTT, True))
      .Add(CommandIndexes.mnuFinCAFProvisionalBatchClaim.ToString, New MenuToolbarCommand("VoucherClaimReport", ControlText.MnuSVoucherClaimReport, CommandIndexes.mnuFinCAFProvisionalBatchClaim, "SCFCPB", , ControlText.MnuSVoucherClaimReportTT, True))
      .Add(CommandIndexes.mnuFinCAFCreateCardSalesReport.ToString, New MenuToolbarCommand("ManualCAFCardSalesClaimReport", ControlText.MnuSManualCAFCardSalesClaimReport, CommandIndexes.mnuFinCAFCreateCardSalesReport, "SCFCCS", , ControlText.MnuSManualCAFCardSalesClaimReportTT, True))
      .Add(CommandIndexes.mnuFinCAFLoadPaymentData.ToString, New MenuToolbarCommand("LoadPaymentData", ControlText.MnuSLoadPaymentData, CommandIndexes.mnuFinCAFLoadPaymentData, "SCFCPL", , ControlText.MnuSLoadPaymentDataTT, True))
      .Add(CommandIndexes.mnuFinCAFReconcilePaymentData.ToString, New MenuToolbarCommand("ReconcilePaymentData", ControlText.MnuSReconcilePaymentData, CommandIndexes.mnuFinCAFReconcilePaymentData, "SCFCPR", , ControlText.MnuSReconcilePaymentDataTT, True))

      .Add(CommandIndexes.mnuCMProcessMetaData.ToString, New MenuToolbarCommand("CMProcessMetaData", ControlText.MnuSCMProcessMetaData, CommandIndexes.mnuCMProcessMetaData, "SCSCMD", , ControlText.MnuSCMProcessMetaData, True))
      .Add(CommandIndexes.mnuCMProcessEventData.ToString, New MenuToolbarCommand("CMProcessEventData", ControlText.MnuSCMProcessEventData, CommandIndexes.mnuCMProcessEventData, "SCSCME", , ControlText.MnuSCMProcessEventData, True))
      .Add(CommandIndexes.mnuCMProcessTotals.ToString, New MenuToolbarCommand("CMProcessTotals", ControlText.MnuSCMProcessTotals, CommandIndexes.mnuCMProcessTotals, "SCSCMT", , ControlText.MnuSCMProcessTotals, True))
      .Add(CommandIndexes.mnuCMUpdateBulkMailer.ToString, New MenuToolbarCommand("CMUpdateBulkMailer", ControlText.MnuSCMUpdateBulkMailer, CommandIndexes.mnuCMUpdateBulkMailer, "SCSBLK", , ControlText.MnuSCMUpdateBulkMailer, True))

      .Add(CommandIndexes.mnuFinCCBatches.ToString, New MenuToolbarCommand("CreateCCCABatches", ControlText.MnuSCreateCCCABatches, CommandIndexes.mnuFinCCBatches, "SCFCCB", , ControlText.MnuSCreateCCCABatchesTT, True))
      .Add(CommandIndexes.mnuFinCCCreateFile.ToString, New MenuToolbarCommand("CCCAClaimFileCreation", ControlText.MnuSCCCAClaimFileCreation, CommandIndexes.mnuFinCCCreateFile, "SCFCCF", , ControlText.MnuSCCCAClaimFileCreationTT, True))
      .Add(CommandIndexes.mnuFinCCCreateReport.ToString, New MenuToolbarCommand("ManualCCCAClaim", ControlText.MnuSManualCCCAClaim, CommandIndexes.mnuFinCCCreateReport, "SCFCMC", , ControlText.MnuSManualCCCAClaimTT, True))
      .Add(CommandIndexes.mnuFinCCCreateCardSalesFile.ToString, New MenuToolbarCommand("CardSalesClaimFileCreation", ControlText.MnuSCardSalesClaimFileCreation, CommandIndexes.mnuFinCCCreateCardSalesFile, "SCFCSF", , ControlText.MnuSCardSalesClaimFileCreationTT, True))
      .Add(CommandIndexes.mnuFinCCCreateCardSalesReport.ToString, New MenuToolbarCommand("ManualCardSalesClaim", ControlText.MnuSManualCardSalesClaim, CommandIndexes.mnuFinCCCreateCardSalesReport, "SCFCMS", , ControlText.MnuSManualCardSalesClaimTT, True))
      .Add(CommandIndexes.mnuFinCCAuthorisationsReport.ToString, New MenuToolbarCommand("AuthorisationsReport", ControlText.MnuSAuthorisationsReport, CommandIndexes.mnuFinCCAuthorisationsReport, "SCFCAR", , ControlText.MnuSAuthorisationsReportTT, True))

      .Add(CommandIndexes.mnuFinCSTransferInvoices.ToString, New MenuToolbarCommand("TransferInvoices", ControlText.MnuSTransferInvoices, CommandIndexes.mnuFinCSTransferInvoices, "SCFMTI", , ControlText.MnuSTransferInvoicesTT, True))
      .Add(CommandIndexes.mnuFinCSTransferCustomers.ToString, New MenuToolbarCommand("TransferCustomers", ControlText.MnuSTransferCustomers, CommandIndexes.mnuFinCSTransferCustomers, "SCFMTC", , ControlText.MnuSTransferCustomersTT, True))
      .Add(CommandIndexes.mnuFinCSStatementGeneration.ToString, New MenuToolbarCommand("StatementGeneration", ControlText.MnuSStatementGeneration, CommandIndexes.mnuFinCSStatementGeneration, "SCFMSG", , ControlText.MnuSStatementGenerationTT, True))

      .Add(CommandIndexes.mnuFinFastDataEntryMaint.ToString, New MenuToolbarCommand("FastDataEntryMaintenance", ControlText.MnuSFDEMaintenance, CommandIndexes.mnuFinFastDataEntryMaint, "SCFFDE"))

      .Add(CommandIndexes.mnuDeDuplicationContactMerge.ToString, New MenuToolbarCommand("ContactMerge", ControlText.MnuSContactMerge, CommandIndexes.mnuDeDuplicationContactMerge, "SCDDCM", , ControlText.MnuSContactMergeTT, True))
      .Add(CommandIndexes.mnuDeDuplicationAddressMerge.ToString, New MenuToolbarCommand("AddressMerge", ControlText.MnuSAddressMerge, CommandIndexes.mnuDeDuplicationAddressMerge, "SCDDAM", , ControlText.MnuSAddressMergeTT, True))
      .Add(CommandIndexes.mnuDeDuplicationOrganisationMerge.ToString, New MenuToolbarCommand("OrganisationMerge", ControlText.MnuSOrganisationMerge, CommandIndexes.mnuDeDuplicationOrganisationMerge, "SCDDOM", , ControlText.MnuSOrganisationMergeTT, True))
      .Add(CommandIndexes.mnuDeDuplicationAmalgamateOrganisations.ToString, New MenuToolbarCommand("AmalgamateOrganisations", ControlText.MnuSAmalgamateOrganisations, CommandIndexes.mnuDeDuplicationAmalgamateOrganisations, "SCDDAO", , ControlText.MnuSAmalgamateOrganisationsTT, True))
      .Add(CommandIndexes.mnuDeDuplicationContactDeDuplication.ToString, New MenuToolbarCommand("ContactDeDuplication", ControlText.MnuSContactDeDuplication, CommandIndexes.mnuDeDuplicationContactDeDuplication, "SCDDCD", , ControlText.MnuSContactDeDuplicationTT, True))
      .Add(CommandIndexes.mnuDeDuplicationBulkAddressMerge.ToString, New MenuToolbarCommand("BulkAddressMerge", ControlText.MnuSBulkAddressMerge, CommandIndexes.mnuDeDuplicationBulkAddressMerge, "SCDDBA", , ControlText.MnuSBulkAddressMergeTT, True))
      .Add(CommandIndexes.mnuDeDuplicationBulkContactMerge.ToString, New MenuToolbarCommand("BulkContactMerge", ControlText.MnuSBulkContactMerge, CommandIndexes.mnuDeDuplicationBulkContactMerge, "SCDDBM", , ControlText.MnuSBulkContactMergeTT, True))
      .Add(CommandIndexes.mnuDeDuplicationBulkOrganisationMerge.ToString, New MenuToolbarCommand("BulkOrganisationMerge", ControlText.MnuSBulkOrganisationMerge, CommandIndexes.mnuDeDuplicationBulkOrganisationMerge, "SCDDBO", , ControlText.MnuSBulkOrganisationMergeTT, True))
      .Add(CommandIndexes.mnuDeDuplicateProcessDuplicateContacts.ToString, New MenuToolbarCommand("ProcessDuplicateContact", ControlText.MnuSProcessDuplicateContact, CommandIndexes.mnuDeDuplicateProcessDuplicateContacts, "SCDDPD", , ControlText.MnuSProcessDuplicateContactTT, True))

      .Add(CommandIndexes.mnuFinDDMandateFile.ToString, New MenuToolbarCommand("MandateFileCreation", ControlText.MnuSMandateFileCreation, CommandIndexes.mnuFinDDMandateFile, "SCFDMF", , ControlText.MnuSMandateFileCreationTT, True))
      .Add(CommandIndexes.mnuFinDDBatches.ToString, New MenuToolbarCommand("CreateDirectDebitBatches", ControlText.MnuSCreateDirectDebitBatches, CommandIndexes.mnuFinDDBatches, "SCFDDB", , ControlText.MnuSCreateDirectDebitBatchesTT, True))
      .Add(CommandIndexes.mnuFinDDClaimFile.ToString, New MenuToolbarCommand("ClaimFileCreation", ControlText.MnuSClaimFileCreation, CommandIndexes.mnuFinDDClaimFile, "SCFDCF", , ControlText.MnuSClaimFileCreationTT, True))
      .Add(CommandIndexes.mnuFinDDUploadBacsMessagingData.ToString, New MenuToolbarCommand("UploadBACSMessagingData", ControlText.MnuSUploadBACSMessagingData, CommandIndexes.mnuFinDDUploadBacsMessagingData, "SCFDCB", , ControlText.MnuSUploadBACSMessagingDataTT, True))
      .Add(CommandIndexes.mnuFinDDDirectCreditFile.ToString, New MenuToolbarCommand("CreditFileCreation", ControlText.MnuSCreditFileCreation, CommandIndexes.mnuFinDDDirectCreditFile, "SCFDRF", , ControlText.MnuSCreditFileCreationTT, True))
      .Add(CommandIndexes.mnuFinDDBACSRejections.ToString, New MenuToolbarCommand("ProcessBACSMessaging", ControlText.MnuSProcessBACSMessaging, CommandIndexes.mnuFinDDBACSRejections, "SCFDBR", , ControlText.MnuSProcessBACSMessagingTT, True))
      .Add(CommandIndexes.mnuFinDDConvertManualDirectDebits.ToString, New MenuToolbarCommand("ConvertManualDirectDebits", ControlText.MnuSConvertManualDirectDebits, CommandIndexes.mnuFinDDConvertManualDirectDebits, "SCFDCM", , ControlText.MnuSConvertManualDirectDebitsTT, True))

      .Add(CommandIndexes.mnuFinDBCreateUnallocatedBoxes.ToString, New MenuToolbarCommand("CreateUnallocatedBoxes", ControlText.MnuSCreateUnallocatedBoxes, CommandIndexes.mnuFinDBCreateUnallocatedBoxes, "SCFDBU", , ControlText.MnuSCreateUnallocatedBoxesTT, True))
      .Add(CommandIndexes.mnuFinDBPrintThankYouLetters.ToString, New MenuToolbarCommand("PrintThankYouLetters", ControlText.MnuSPrintThankYouLetters, CommandIndexes.mnuFinDBPrintThankYouLetters, "SCFDST", , ControlText.MnuSPrintThankYouLettersTT, True))
      .Add(CommandIndexes.mnuFinDBPrintAdviceNotes.ToString, New MenuToolbarCommand("PrintAdviceNotes", ControlText.MnuSPrintAdviceNotes, CommandIndexes.mnuFinDBPrintAdviceNotes, "SCFDSN", , ControlText.MnuSPrintAdviceNotesTT, True))
      .Add(CommandIndexes.mnuFinDBPrintPackingSlips.ToString, New MenuToolbarCommand("PrintPackingSlips", ControlText.MnuSPrintPackingSlips, CommandIndexes.mnuFinDBPrintPackingSlips, "SCFDSP", , ControlText.MnuSPrintPackingSlipsTT, True))
      .Add(CommandIndexes.mnuFinDBPrintBoxLabels.ToString, New MenuToolbarCommand("PrintBoxLabels", ControlText.MnuSPrintBoxLabels, CommandIndexes.mnuFinDBPrintBoxLabels, "SCFDSL", , ControlText.MnuSPrintBoxLabelsTT, True))
      .Add(CommandIndexes.mnuFinDBSetShippingInformation.ToString, New MenuToolbarCommand("SetShippingInformation", ControlText.MnuSSetShippingInformation, CommandIndexes.mnuFinDBSetShippingInformation, "SCFDSB", , ControlText.MnuSSetShippingInformationTT, True))
      .Add(CommandIndexes.mnuFinDBSetArrivalInformation.ToString, New MenuToolbarCommand("SetArrivalInformation", ControlText.MnuSSetArrivalInformation, CommandIndexes.mnuFinDBSetArrivalInformation, "SCFDSA", , ControlText.MnuSSetArrivalInformationTT, True))

      .Add(CommandIndexes.mnuFinDBRepOpenBoxes.ToString, New MenuToolbarCommand("OpenBoxes", ControlText.MnuSOpenBoxes, CommandIndexes.mnuFinDBRepOpenBoxes, "SCFDRO", , ControlText.MnuSOpenBoxesTT, True))
      .Add(CommandIndexes.mnuFinDBRepUnAllocatedDonations.ToString, New MenuToolbarCommand("UnallocatedDonations", ControlText.MnuSUnallocatedDonations, CommandIndexes.mnuFinDBRepUnAllocatedDonations, "SCFDRU", , ControlText.MnuSUnallocatedDonationsTT, True))
      .Add(CommandIndexes.mnuFinDBRepAllocatedDonations.ToString, New MenuToolbarCommand("AllocatedDonations", ControlText.MnuSAllocatedDonations, CommandIndexes.mnuFinDBRepAllocatedDonations, "SCFDRA", , ControlText.MnuSAllocatedDonationsTT, True))
      .Add(CommandIndexes.mnuFinDBRepDonorDetails.ToString, New MenuToolbarCommand("DonorDetails", ControlText.MnuSDonorDetails, CommandIndexes.mnuFinDBRepDonorDetails, "SCFDRD", , ControlText.MnuSDonorDetailsTT, True))
      .Add(CommandIndexes.mnuFinDBRepClosedByLocation.ToString, New MenuToolbarCommand("ClosedBoxesByLocation", ControlText.MnuSClosedBoxesByLocation, CommandIndexes.mnuFinDBRepClosedByLocation, "SCFDRC", , ControlText.MnuSClosedBoxesByLocationTT, True))
      .Add(CommandIndexes.mnuFinDBRepRollOfHonour.ToString, New MenuToolbarCommand("GenerateRollOfHonour", ControlText.MnuSGenerateRollOfHonour, CommandIndexes.mnuFinDBRepRollOfHonour, "SCFDRR", , ControlText.MnuSGenerateRollOfHonourTT, True))

      .Add(CommandIndexes.mnuDutchLoadPayments.ToString, New MenuToolbarCommand("DutchLoadPayments", ControlText.MnuSDutchLoadPayments, CommandIndexes.mnuDutchLoadPayments, "SCFDPL", , ControlText.MnuSDutchLoadPaymentsTT, True))
      .Add(CommandIndexes.mnuDutchProcessPayments.ToString, New MenuToolbarCommand("DutchProcessPayments", ControlText.MnuSDutchProcessPayments, CommandIndexes.mnuDutchProcessPayments, "SCFDPP", , ControlText.MnuSDutchProcessPaymentsTT, True))

      .Add(CommandIndexes.mnuFinGADConfirmation.ToString, New MenuToolbarCommand("DeclarationConfirmation", ControlText.MnuSDeclarationConfirmation, CommandIndexes.mnuFinGADConfirmation, "SCFGDC", , ControlText.MnuSDeclarationConfirmationTT, True))
      .Add(CommandIndexes.mnuFinGADPotentialClaim.ToString, New MenuToolbarCommand("CreatePotentialClaim", ControlText.MnuSGADCreatePotentialClaim, CommandIndexes.mnuFinGADPotentialClaim, "SCFGPC", , ControlText.MnuSGADCreatePotentialClaimTT, True))
      .Add(CommandIndexes.mnuFinGADTaxClaim.ToString, New MenuToolbarCommand("CreateTaxClaim", ControlText.MnuSGADCreateTaxClaim, CommandIndexes.mnuFinGADTaxClaim, "SCFGCT", , ControlText.MnuSGADCreateTaxClaimTT, True))
      .Add(CommandIndexes.mnuFinGADGiftDetails.ToString, New MenuToolbarCommand("ReprintClaimDetailsReport", ControlText.MnuSGADReprintClaimDetailsReport, CommandIndexes.mnuFinGADGiftDetails, "SCFGCD", , ControlText.MnuSGADReprintClaimDetailsReportTT, True))
      .Add(CommandIndexes.mnuFinGADGiftAnalysis.ToString, New MenuToolbarCommand("ClaimAnalysisReport", ControlText.MnuSClaimAnalysisReport, CommandIndexes.mnuFinGADGiftAnalysis, "SCFGCA", , ControlText.MnuSClaimAnalysisReportTT, True))
      .Add(CommandIndexes.mnuFinBulkGiftAidUpdate.ToString, New MenuToolbarCommand("BulkGiftAidUpdate", ControlText.MnuSBulkGiftAidUpdate, CommandIndexes.mnuFinBulkGiftAidUpdate, "SCFGBU", , ControlText.MnuSBulkGiftAidUpdateTT, True))

      .Add(CommandIndexes.mnuFinGASPotentialClaim.ToString, New MenuToolbarCommand("CreatePotentialClaim", ControlText.MnuSGASCreatePotentialClaim, CommandIndexes.mnuFinGASPotentialClaim, "SCFGSP", , ControlText.MnuSGASCreatePotentialClaimTT, True))
      .Add(CommandIndexes.mnuFinGASTaxClaim.ToString, New MenuToolbarCommand("CreateTaxClaim", ControlText.MnuSGASCreateTaxClaim, CommandIndexes.mnuFinGASTaxClaim, "SCFGST", , ControlText.MnuSGASCreateTaxClaimTT, True))
      .Add(CommandIndexes.mnuFinGASClaimDetails.ToString, New MenuToolbarCommand("ReprintClaimDetailsReport", ControlText.MnuSGASReprintClaimDetailsReport, CommandIndexes.mnuFinGASClaimDetails, "SCFGSC", , ControlText.MnuSGASReprintClaimDetailsReportTT, True))

      .Add(CommandIndexes.mnuFinGAIPotentialClaim.ToString, New MenuToolbarCommand("IrishCreatePotentialClaim", ControlText.MnuSGAICreatePotentialClaim, CommandIndexes.mnuFinGAIPotentialClaim, "SCFGIP", , ControlText.MnuSGAICreatePotentialClaimTT, True))
      .Add(CommandIndexes.mnuFinGAITaxClaim.ToString, New MenuToolbarCommand("IrishCreateTaxClaim", ControlText.MnuSGAICreateTaxClaim, CommandIndexes.mnuFinGAITaxClaim, "SCFGIT", , ControlText.MnuSGAICreateTaxClaimTT, True))
      .Add(CommandIndexes.mnuFinGAIClaimDetails.ToString, New MenuToolbarCommand("IrishReprintClaimDetailsReport", ControlText.MnuSGAIReprintClaimDetailsReport, CommandIndexes.mnuFinGAIClaimDetails, "SCFGIR", , ControlText.MnuSGAIReprintClaimDetailsReportTT, True))

      .Add(CommandIndexes.mnuFinancialIncentives.ToString, New MenuToolbarCommand("Maintain", ControlText.MnuSMaintain, CommandIndexes.mnuFinancialIncentives, "SCFIMI", , ControlText.MnuSMaintainTT, True))

      .Add(CommandIndexes.mnuMailingDocsProduce.ToString, New MenuToolbarCommand("ProduceMailingDocuments", ControlText.MnuSProduceMailingDocuments, CommandIndexes.mnuMailingDocsProduce, "SCMMRE", , ControlText.MnuSProduceMailingDocuments, True))
      .Add(CommandIndexes.mnuMailingDocsFind.ToString, New MenuToolbarCommand("FindMailingDocuments", ControlText.MnuSFindMailingDocuments, CommandIndexes.mnuMailingDocsFind, "SCMMFD", , ControlText.MnuSFindMailingDocuments, True))
      .Add(CommandIndexes.mnuMailingTYL.ToString, New MenuToolbarCommand("ThankYouLetters", ControlText.MnuSmailingTYL, CommandIndexes.mnuMailingTYL, "SCMMTL", , ControlText.MnuSmailingTYL, True))
      .Add(CommandIndexes.mnuMailingFinder.ToString, New MenuToolbarCommand("FindMailings", ControlText.MnuSMailingFinder, CommandIndexes.mnuMailingFinder, "SCMMFM", , ControlText.MnuSMailingFinder, True))
      .Add(CommandIndexes.mnuEMailProcessor.ToString, New MenuToolbarCommand("EMailProcessor", ControlText.MnuSEmailProcessor, CommandIndexes.mnuEMailProcessor, "SCMMEP", , ControlText.MnuSEmailProcessor, True))
      .Add(CommandIndexes.mnuMailingListAllContacts.ToString, New MenuToolbarCommand("ListAllContacts", ControlText.MnuSMailingListAllContacts, CommandIndexes.mnuMailingListAllContacts, "SCMMAC", , ControlText.MnuSMailingListAllContacts, True))

      .Add(CommandIndexes.mnuMktGenerateData.ToString, New MenuToolbarCommand("MarketingGenerateData", ControlText.MnuSMarketingGenerateData, CommandIndexes.mnuMktGenerateData, "SCMKGD", , ControlText.MnuSMarketingGenerateDataTT, True))

      .Add(CommandIndexes.mnuMemFutureChanges.ToString, New MenuToolbarCommand("MemFutureChanges", ControlText.MnuSMemFutureChanges, CommandIndexes.mnuMemFutureChanges, "SCMEFM", , ControlText.MnuSMemFutureChangesTT, True))
      .Add(CommandIndexes.mnuMemCards.ToString, New MenuToolbarCommand("MemCards", ControlText.MnuSMemCards, CommandIndexes.mnuMemCards, "SCMEMC", , ControlText.MnuSMemCardsTT, ))
      .Add(CommandIndexes.mnuMemSuspension.ToString, New MenuToolbarCommand("MemSuspension", ControlText.MnuSMemSuspension, CommandIndexes.mnuMemSuspension, "SCMEMS", , ControlText.MnuSMemSuspensionTT, True))
      .Add(CommandIndexes.mnuMemFulfilment.ToString, New MenuToolbarCommand("MemFulfillment", ControlText.MnuSMemFulfilment, CommandIndexes.mnuMemFulfilment, "SCMENM", , ControlText.MnuSMemFulfilmentTT, True))

      .Add(CommandIndexes.mnuMemBranchDonations.ToString, New MenuToolbarCommand("MemBranchDonations", ControlText.MnuSMemBranchDonations, CommandIndexes.mnuMemBranchDonations, "SCMRBD", , ControlText.MnuSMemBranchDonationsTT, True))
      .Add(CommandIndexes.mnuMemBranchIncome.ToString, New MenuToolbarCommand("MemBranchIncome", ControlText.MnuSMemBranchIncome, CommandIndexes.mnuMemBranchIncome, "SCMRBI", , ControlText.MnuSMemBranchIncomeTT, True))
      .Add(CommandIndexes.mnuMemJuniorAnalysis.ToString, New MenuToolbarCommand("MemJuniorAnalysis", ControlText.MnuSMemJuniorAnalysis, CommandIndexes.mnuMemJuniorAnalysis, "SCMRJA", , ControlText.MnuSMemJuniorAnalysisTT, True))
      .Add(CommandIndexes.mnuMemAssumedVotingRights.ToString, New MenuToolbarCommand("MemAssumedVotingRights", ControlText.MnuSMemAssumedVotingRights, CommandIndexes.mnuMemAssumedVotingRights, "SCMRAV", , ControlText.MnuSMemAssumedVotingRightsTT, True))
      .Add(CommandIndexes.mnuMemBallotPaperProduction.ToString, New MenuToolbarCommand("MemBallotPaperProduction", ControlText.MnuSMemBallotPaperProduction, CommandIndexes.mnuMemBallotPaperProduction, "SCMRBP", , ControlText.MnuSMemBallotPaperProductionTT, True))

      .Add(CommandIndexes.mnuMemGenerateStatistics.ToString, New MenuToolbarCommand("MemGenerateStatistics", ControlText.MnuSMemGenerateStatistics, CommandIndexes.mnuMemGenerateStatistics, "SCMSGD", , ControlText.MnuSMemGenerateStatisticsTT, True))
      .Add(CommandIndexes.mnuMemStatisticsDetailed.ToString, New MenuToolbarCommand("MemStatisticsDetailed", ControlText.MnuSMemStatisticsDetailed, CommandIndexes.mnuMemStatisticsDetailed, "SCMSDR", , ControlText.MnuSMemStatisticsDetailedTT, True))
      .Add(CommandIndexes.mnuMemStatisticsSummary.ToString, New MenuToolbarCommand("MemStatisticsSummary", ControlText.MnuSMemStatisticsSummary, CommandIndexes.mnuMemStatisticsSummary, "SCMSSR", , ControlText.MnuSMemStatisticsSummaryTT, True))

      .Add(CommandIndexes.mnuFinNomSummaryReport.ToString, New MenuToolbarCommand("SummaryReport", ControlText.MnuSNomSummaryReport, CommandIndexes.mnuFinNomSummaryReport, "SCFNSR", , ControlText.MnuSNomSummaryReportTT, True))
      .Add(CommandIndexes.mnuFinNomDetailReport.ToString, New MenuToolbarCommand("DetailedReport", ControlText.MnuSNomDetailedReport, CommandIndexes.mnuFinNomDetailReport, "SCFNDR", , ControlText.MnuSNomDetailedReportTT, True))

      .Add(CommandIndexes.mnuFinPISLoadStatement.ToString, New MenuToolbarCommand("LoadBankStatementData", ControlText.MnuSPISLoadBankStatementData, CommandIndexes.mnuFinPISLoadStatement, "SCFPCL", , ControlText.MnuSPISLoadBankStatementDataTT, True))
      .Add(CommandIndexes.mnuFinPISReconciliation.ToString, New MenuToolbarCommand("AutomatedReconciliation", ControlText.MnuSPISAutomatedReconciliation, CommandIndexes.mnuFinPISReconciliation, "SCFPCR", , ControlText.MnuSPISAutomatedReconciliationTT, True))

      .Add(CommandIndexes.mnuFinPPRenewals.ToString, New MenuToolbarCommand("RenewalsReminders", ControlText.MnuSRenewalsReminders, CommandIndexes.mnuFinPPRenewals, "SCFPRR", , ControlText.MnuSRenewalsRemindersTT, True))
      .Add(CommandIndexes.mnuFinPPRemoveArrears.ToString, New MenuToolbarCommand("RemoveOldDetailsArrears", ControlText.MnuSRemoveOldDetailsArrears, CommandIndexes.mnuFinPPRemoveArrears, "SCFPRO", , ControlText.MnuSRemoveOldDetailsArrearsTT, True))
      .Add(CommandIndexes.mnuFinPPExpiry.ToString, New MenuToolbarCommand("CancelExpiredPaymentPlans", ControlText.MnuSCancelExpiredPaymentPlans, CommandIndexes.mnuFinPPExpiry, "SCFPCE", , ControlText.MnuSCancelExpiredPaymentPlansTT, True))
      .Add(CommandIndexes.mnuFinPPNonMemberFulfilment.ToString, New MenuToolbarCommand("Non-memberFulfilment", ControlText.MnuSNonMemberFulfilment, CommandIndexes.mnuFinPPNonMemberFulfilment, "SCFPNM", , ControlText.MnuSNonMemberFulfilmentTT, True))
      .Add(CommandIndexes.mnuFinPPUpdateProducts.ToString, New MenuToolbarCommand("UpdateProducts", ControlText.MnuSUpdateProducts, CommandIndexes.mnuFinPPUpdateProducts, "SCFPCP", , ControlText.MnuSUpdateProductsTT, True))
      .Add(CommandIndexes.mnuFinPPApplySurcharges.ToString, New MenuToolbarCommand("ApplySurcharges", ControlText.MnuSApplySurcharges, CommandIndexes.mnuFinPPApplySurcharges, "SCFPOS", , ControlText.MnuSApplySurchargesTT, True))
      .Add(CommandIndexes.mnuFinPPReCalcLoanInterest.ToString, New MenuToolbarCommand("Re-calculate Loan Interest", ControlText.MnuSRecalcLoanInterest, CommandIndexes.mnuFinPPReCalcLoanInterest, "SCFPLI", , ControlText.MnuSRecalcLoanInterest, True))
      .Add(CommandIndexes.mnuFinPPUpdateLoanInterestRates.ToString, New MenuToolbarCommand("Update Loan Interest Rates", ControlText.MnuSUpdateLoanInterestRates, CommandIndexes.mnuFinPPUpdateLoanInterestRates, "SCFPUL", , ControlText.MnuSUpdateLoanInterestRatesTT, True))
      .Add(CommandIndexes.mnuFinPPTransferPaymentPlanChanges.ToString, New MenuToolbarCommand("TransferPaymentPlanChanges", ControlText.MnuSTransferPaymentPlanChanges, CommandIndexes.mnuFinPPTransferPaymentPlanChanges, "SCFEPP", , ControlText.MnuSTransferPaymentPlanChangesTT, True))

      .Add(CommandIndexes.mnuFinGAYELoadPayments.ToString, New MenuToolbarCommand("LoadPreTaxPaymentData", ControlText.MnuSLoadPreTaxPaymentData, CommandIndexes.mnuFinGAYELoadPayments, "SCFYPL", , ControlText.MnuSLoadPreTaxPaymentDataTT, True))
      .Add(CommandIndexes.mnuFinGAYEReconciliation.ToString, New MenuToolbarCommand("PreTaxAutomatedReconciliation", ControlText.MnuSPreTaxAutomatedReconciliation, CommandIndexes.mnuFinGAYEReconciliation, "SCFYRC", , ControlText.MnuSPreTaxAutomatedReconciliationTT, True))
      .Add(CommandIndexes.mnuFinGAYEBulkCancellation.ToString, New MenuToolbarCommand("PreTaxPledgeBulkCancellation", ControlText.MnuSPreTaxPledgeBulkCancellation, CommandIndexes.mnuFinGAYEBulkCancellation, "SCFYBC", , ControlText.MnuSPreTaxPledgeBulkCancellationTT, True))
      .Add(CommandIndexes.mnuFinGAYEPostTaxPGReconciliation.ToString, New MenuToolbarCommand("PostTaxAutomatedReconciliation", ControlText.MnuSPostTaxAutomatedReconciliation, CommandIndexes.mnuFinGAYEPostTaxPGReconciliation, "SCFYPT", , ControlText.MnuSPostTaxAutomatedReconciliationTT, True))

      .Add(CommandIndexes.mnuFinPOTransPayments.ToString, New MenuToolbarCommand("TransferPayments", ControlText.MnuSTransferPayments, CommandIndexes.mnuFinPOTransPayments, "SCFPPP", , ControlText.MnuSTransferPaymentsTT, True))
      .Add(CommandIndexes.mnuFinPOTransSuppliers.ToString, New MenuToolbarCommand("TransferSuppliers", ControlText.MnuSTransferSuppliers, CommandIndexes.mnuFinPOTransSuppliers, "SCFPTS", , ControlText.MnuSTransferSuppliersTT, True))
      .Add(CommandIndexes.mnuFinPOAuthorisePayments.ToString, New MenuToolbarCommand("AuthorisePayments", ControlText.MnuSAuthorisePayments, CommandIndexes.mnuFinPOAuthorisePayments, "SCFPRA", , ControlText.MnuSAuthorisePaymentsTT, True))
      .Add(CommandIndexes.mnuFinPOProcessPayments.ToString, New MenuToolbarCommand("ProcessPayments", ControlText.MnuSPOProcessPayments, CommandIndexes.mnuFinPOProcessPayments, "SCFPPS", , ControlText.MnuSPOProcessPaymentsTT, True))
      .Add(CommandIndexes.mnuFinPOAutoGenerate.ToString, New MenuToolbarCommand("AutoGenerate", ControlText.MnuSAutoGenerate, CommandIndexes.mnuFinPOAutoGenerate, "SCFPAG", , ControlText.MnuSAutoGenerateTT, True))
      .Add(CommandIndexes.mnuFinPOPrint.ToString, New MenuToolbarCommand("Print", ControlText.MnuSPrint, CommandIndexes.mnuFinPOPrint, "SCFPPO", , ControlText.MnuSPrint, True))
      .Add(CommandIndexes.mnuFinChequeProduction.ToString, New MenuToolbarCommand("Payment Production", ControlText.MnuSChequeProduction, CommandIndexes.mnuFinChequeProduction, "SCFPQP", , ControlText.MnuSChequeProduction, True))

      .Add(CommandIndexes.mnuMailingMembers.ToString, New MenuToolbarCommand("Members", ControlText.MnuSMailingMembers, CommandIndexes.mnuMailingMembers, "SCMMMB", , ControlText.MnuSMailingMembers))
      .Add(CommandIndexes.mnuMailingStandingOrders.ToString, New MenuToolbarCommand("StandingOrders", ControlText.MnuSMailingStandingOrders, CommandIndexes.mnuMailingStandingOrders, "SCMMSO", , ControlText.MnuSMailingStandingOrders))
      .Add(CommandIndexes.mnuMailingDirectDebit.ToString, New MenuToolbarCommand("DirectDebit", ControlText.MnuSMailingDirectDebit, CommandIndexes.mnuMailingDirectDebit, "SCMMDB", , ControlText.MnuSMailingDirectDebit))
      .Add(CommandIndexes.mnuMailingPayers.ToString, New MenuToolbarCommand("Payers", ControlText.MnuSMailingPayers, CommandIndexes.mnuMailingPayers, "SCMMPA", , ControlText.MnuSMailingPayers))
      .Add(CommandIndexes.mnuMailingSubscriptions.ToString, New MenuToolbarCommand("Subscriptions", ControlText.MnuSMailingSubscriptions, CommandIndexes.mnuMailingSubscriptions, "SCMMSC", , ControlText.MnuSMailingSubscriptions))
      .Add(CommandIndexes.mnuMailingSelectionTester.ToString, New MenuToolbarCommand("SelectionTester", ControlText.MnuSMailingSelectionTester, CommandIndexes.mnuMailingSelectionTester, "SCMMST", , ControlText.MnuSMailingSelectionTester))

      .Add(CommandIndexes.mnuMailingEventsBookings.ToString, New MenuToolbarCommand("EventBookings", ControlText.MnuSMailingEventBookings, CommandIndexes.mnuMailingEventsBookings, "SCMMEB", , ControlText.MnuSMailingEventBookings))
      .Add(CommandIndexes.mnuMailingEventsDelegates.ToString, New MenuToolbarCommand("EventDelegates", ControlText.MnuSMailingEventDelegates, CommandIndexes.mnuMailingEventsDelegates, "SCMMED", , ControlText.MnuSMailingEventDelegates))
      .Add(CommandIndexes.mnuMailingEventsPersonnel.ToString, New MenuToolbarCommand("EventPersonnel", ControlText.MnuSMailingEventPersonnel, CommandIndexes.mnuMailingEventsPersonnel, "SCMMEN", , ControlText.MnuSMailingEventPersonnel))
      .Add(CommandIndexes.mnuMailingEventsSponsors.ToString, New MenuToolbarCommand("EventSponsors", ControlText.MnuSMailingEventSponsors, CommandIndexes.mnuMailingEventsSponsors, "SCMMES", , ControlText.MnuSMailingEventSponsors))

      .Add(CommandIndexes.mnuMailingExamsBookings.ToString, New MenuToolbarCommand("ExamBookings", ControlText.MnuSMailingExamBookings, CommandIndexes.mnuMailingExamsBookings, "SCMMXB", , ControlText.MnuSMailingExamBookings))
      .Add(CommandIndexes.mnuMailingExamsCandidates.ToString, New MenuToolbarCommand("ExamCandidates", ControlText.MnuSMailingExamCandidates, CommandIndexes.mnuMailingExamsCandidates, "SCMMXC", , ControlText.MnuSMailingExamCandidates))

      .Add(CommandIndexes.mnuMailingPreTextPayrollGivingPledges.ToString, New MenuToolbarCommand("PreTextPayrollGivingPledges", ControlText.MnuSMailingPreTextPayrollGivingPledges, CommandIndexes.mnuMailingPreTextPayrollGivingPledges, "SCMMPG", , ControlText.MnuSMailingPreTextPayrollGivingPledges))
      .Add(CommandIndexes.mnuMailingSelectionManager.ToString, New MenuToolbarCommand("SelectionManager", ControlText.MnuSMailingSelectionManager, CommandIndexes.mnuMailingSelectionManager, "SCMMSM", , ControlText.MnuSMailingSelectionManager))
      .Add(CommandIndexes.mnuMailingIrishGiftAid.ToString, New MenuToolbarCommand("IrishGiftAid", ControlText.MnuSMailingIrishGiftAid, CommandIndexes.mnuMailingIrishGiftAid, "SCMMGA", , ControlText.MnuSMailingIrishGiftAid))
      .Add(CommandIndexes.mnuFinPRPriceChange.ToString, New MenuToolbarCommand("PriceChangeUpdate", ControlText.MnuSPriceChangeUpdate, CommandIndexes.mnuFinPRPriceChange, "SCFPPC", , ControlText.MnuSPriceChangeUpdateTT, True))
      .Add(CommandIndexes.mnuFinPRPurchasedProductReport.ToString, New MenuToolbarCommand("PurchasedProductReport", ControlText.MnuSPurchasedProductReport, CommandIndexes.mnuFinPRPurchasedProductReport, "SCFPPR", , ControlText.MnuSPurchasedProductReportTT, True))

      .Add(CommandIndexes.mnuFinSOLoadStatement.ToString, New MenuToolbarCommand("LoadBankStatementData", ControlText.MnuSLoadBankStatementData, CommandIndexes.mnuFinSOLoadStatement, "SCFSLB", , ControlText.MnuSLoadBankStatementDataTT, True))
      .Add(CommandIndexes.mnuFinSOReconciliation.ToString, New MenuToolbarCommand("AutomatedReconciliation", ControlText.MnuSAutomatedReconciliation, CommandIndexes.mnuFinSOReconciliation, "SCFSAR", , ControlText.MnuSAutomatedReconciliationTT, True))
      .Add(CommandIndexes.mnuFinSOCancellation.ToString, New MenuToolbarCommand("BulkCancellation", ControlText.MnuSBulkCancellation, CommandIndexes.mnuFinSOCancellation, "SCFSBC", , ControlText.MnuSBulkCancellationTT, True))
      .Add(CommandIndexes.mnuFinSOManualRec.ToString, New MenuToolbarCommand("ManualReconciliation", ControlText.MnuSManualReconciliation, CommandIndexes.mnuFinSOManualRec, "SCFSMR", , ControlText.MnuSManualReconciliationTT, True))
      .Add(CommandIndexes.mnuFinSOReconReport.ToString, New MenuToolbarCommand("ReconciliationReport", ControlText.MnuSReconciliationReport, CommandIndexes.mnuFinSOReconReport, "SCFSRR", , ControlText.MnuSReconciliationReportTT, True))
      .Add(CommandIndexes.mnuFinSOBankTransactionsReport.ToString, New MenuToolbarCommand("BankTransactionsReport", ControlText.MnuSBankTransactionsReport, CommandIndexes.mnuFinSOBankTransactionsReport, "SCFSBT", , ControlText.MnuSBankTransactionsReportTT, True))

      .Add(CommandIndexes.mnuFinSTPickingLists.ToString, New MenuToolbarCommand("PickingListProduction", ControlText.MnuSPickingListProduction, CommandIndexes.mnuFinSTPickingLists, "SCFSPL", , ControlText.MnuSPickingListProductionTT, True))
      .Add(CommandIndexes.mnuFinSTConfirmAllocation.ToString, New MenuToolbarCommand("ConfirmStockAllocation", ControlText.MnuSConfirmStockAllocation, CommandIndexes.mnuFinSTConfirmAllocation, "SCFSCA", , ControlText.MnuSConfirmStockAllocationTT, True))
      .Add(CommandIndexes.mnuFinSTAllocateToBO.ToString, New MenuToolbarCommand("AllocateStocktoBackOrders", ControlText.MnuSAllocateStocktoBackOrders, CommandIndexes.mnuFinSTAllocateToBO, "SCFSSB", , ControlText.MnuSAllocateStocktoBackOrdersTT, True))
      .Add(CommandIndexes.mnuFinSTDespatchNotes.ToString, New MenuToolbarCommand("DespatchNotes", ControlText.MnuSDespatchNotes, CommandIndexes.mnuFinSTDespatchNotes, "SCFSDN", , ControlText.MnuSDespatchNotesTT, True))
      .Add(CommandIndexes.mnuFinSTDespatchTracking.ToString, New MenuToolbarCommand("DespatchTracking", ControlText.MnuSDespatchTracking, CommandIndexes.mnuFinSTDespatchTracking, "SCFSDT", , ControlText.MnuSDespatchTrackingTT, True))
      .Add(CommandIndexes.mnuFinSTBackOrdersReport.ToString, New MenuToolbarCommand("BackOrdersReport", ControlText.MnuSBackOrdersReport, CommandIndexes.mnuFinSTBackOrdersReport, "SCFSBO", , ControlText.MnuSBackOrdersReportTT, True))
      .Add(CommandIndexes.mnuFinSTSalesAnalysis.ToString, New MenuToolbarCommand("SalesAnalysisMYL", ControlText.MnuSSalesAnalysisMYL, CommandIndexes.mnuFinSTSalesAnalysis, "SCFSSA", , ControlText.MnuSSalesAnalysisMYLTT, True))
      .Add(CommandIndexes.mnuFinSTSalesAnalysisDetailed.ToString, New MenuToolbarCommand("SalesAnalysisDetailed", ControlText.MnuSSalesAnalysisDetailed, CommandIndexes.mnuFinSTSalesAnalysisDetailed, "SCFSSD", , ControlText.MnuSSalesAnalysisDetailedTT, True))
      .Add(CommandIndexes.mnuFinSTSalesAnalysisSummary.ToString, New MenuToolbarCommand("SalesAnalysisSummary", ControlText.MnuSSalesAnalysisSummary, CommandIndexes.mnuFinSTSalesAnalysisSummary, "SCFSSS", , ControlText.MnuSSalesAnalysisSummaryTT, True))
      .Add(CommandIndexes.mnuFinSTMovement.ToString, New MenuToolbarCommand("StockMovement", ControlText.MnuSStockMovement, CommandIndexes.mnuFinSTMovement, "SCFSSM", , ControlText.MnuSStockMovementTT, True))
      .Add(CommandIndexes.mnuFinSTTransferStockToPack.ToString, New MenuToolbarCommand("TransferStockToPack", ControlText.MnuSTransferStockToPack, CommandIndexes.mnuFinSTTransferStockToPack, "SCFSTP", , ControlText.MnuSTransferStockToPackTT, True))
      .Add(CommandIndexes.mnuFinSTExport.ToString, New MenuToolbarCommand("ExportStock", ControlText.MnuSExportStock, CommandIndexes.mnuFinSTExport, "SCFSES", , ControlText.MnuSExportStockTT, True))
      .Add(CommandIndexes.mnuFinSTPurgeBackOrders.ToString, New MenuToolbarCommand("PurgeOldBackOrders", ControlText.MnuSPurgeOldBackOrders, CommandIndexes.mnuFinSTPurgeBackOrders, "SCFSPB", , ControlText.MnuSPurgeOldBackOrdersTT, True))
      .Add(CommandIndexes.mnuFinSTPurgePickingAndDespatch.ToString, New MenuToolbarCommand("PurgeOldPickingAndDespatchData", ControlText.MnuSPurgeOldPickingAndDespatchData, CommandIndexes.mnuFinSTPurgePickingAndDespatch, "SCFSPP", , ControlText.MnuSPurgeOldPickingAndDespatchDataTT, True))
      .Add(CommandIndexes.mnuFinSTPLAwaitingConfirm.ToString, New MenuToolbarCommand("PickingListsAwaitingConfirmation", ControlText.MnuSPickingListsAwaitingConfirmation, CommandIndexes.mnuFinSTPLAwaitingConfirm, "SCFSAC", , ControlText.MnuSPickingListsAwaitingConfirmationTT, True))
      .Add(CommandIndexes.mnuFinSTValuationReport.ToString, New MenuToolbarCommand("StockValuationReport", ControlText.MnuSStockValuationReport, CommandIndexes.mnuFinSTValuationReport, "SCFSSV", , ControlText.MnuSStockValuationReportTT, True))

      .Add(CommandIndexes.mnuAdminAmendmentHistory.ToString, New MenuToolbarCommand("AmendmentHistory", ControlText.MnuSAmendmentHistory, CommandIndexes.mnuAdminAmendmentHistory, "SCAMAH", , ControlText.MnuSAmendmentHistoryTT, True))
      .Add(CommandIndexes.mnuAdminProcessAddressChanges.ToString, New MenuToolbarCommand("ProcessAddressChanges", ControlText.MnuSProcessAddressChanges, CommandIndexes.mnuAdminProcessAddressChanges, "SCAMPA", , ControlText.MnuSProcessAddressChangesTT, True))
      .Add(CommandIndexes.mnuAdminUpdateRegionalData.ToString, New MenuToolbarCommand("UpdateRegionalData", ControlText.MnusUpdateRegionalData, CommandIndexes.mnuAdminUpdateRegionalData, "SCAMRD", , ControlText.MnusUpdateRegionalDataTT, True))
      .Add(CommandIndexes.mnuAdminSetPostDatedContacts.ToString, New MenuToolbarCommand("SetPostDatedContacts", ControlText.MnuSSetPostDatedContacts, CommandIndexes.mnuAdminSetPostDatedContacts, "SCAMPC", , ControlText.MnuSSetPostDatedContactsTT, True))
      .Add(CommandIndexes.mnuAdminUpdateActionStatuses.ToString, New MenuToolbarCommand("UpdateActionStatuses", ControlText.MnuSUpdateActionStatuses, CommandIndexes.mnuAdminUpdateActionStatuses, "SCAMUA", , ControlText.MnuSUpdateActionStatusesTT, True))
      .Add(CommandIndexes.mnuAdminUpdateSearchNames.ToString, New MenuToolbarCommand("UpdateSearchNames", ControlText.MnusUpdateSearchNames, CommandIndexes.mnuAdminUpdateSearchNames, "SCAMSN", , ControlText.MnusUpdateSearchNamesTT, True))
      .Add(CommandIndexes.mnuAdminUpdateMailsort.ToString, New MenuToolbarCommand("UpdateMailsort", ControlText.MnusUpdateMailsort, CommandIndexes.mnuAdminUpdateMailsort, "SCAMUM", , ControlText.MnusUpdateMailsortTT, True))
      .Add(CommandIndexes.mnuAdminUpdatePrincipalUser.ToString, New MenuToolbarCommand("UpdatePrincipalUser", ControlText.MnusUpdatePrincipalUser, CommandIndexes.mnuAdminUpdatePrincipalUser, "SCAUPU", , ControlText.MnusUpdatePrincipalUserTT, True))
      .Add(CommandIndexes.mnuAdminPostcodeValidation.ToString, New MenuToolbarCommand("PostcodeValidation", ControlText.MnuSPostcodeValidation, CommandIndexes.mnuAdminPostcodeValidation, "SCAMPV", , ControlText.MnuSPostcodeValidationTT, True))
      .Add(CommandIndexes.mnuAdminPurgeStickyNotes.ToString, New MenuToolbarCommand("PurgeStickyNotes", ControlText.MnuSPurgeStickyNotes, CommandIndexes.mnuAdminPurgeStickyNotes, "SCAMPS", , ControlText.MnuSPurgeStickyNotesTT, True))

      .Add(CommandIndexes.mnuEventBlockBooking.ToString, New MenuToolbarCommand("EventBlockBooking", ControlText.MnuSEventBlockBooking, CommandIndexes.mnuEventBlockBooking, "SCEVBB", , ControlText.MnuSEventBlockBooking, True))

      .Add(CommandIndexes.mnuExamsMaintenance.ToString, New MenuToolbarCommand("ExamMaintenance", ControlText.MnuSExamMaintenance, CommandIndexes.mnuExamsMaintenance, "SCEXMA", mvImageList16.Images(CommandIndexes.mnuExamsMaintenance), ControlText.MnuSExamMaintenanceTT, True))
      .Add(CommandIndexes.mnuExamEnterResults.ToString, New MenuToolbarCommand("ExamEnterResults", ControlText.MnuSExamEnterResults, CommandIndexes.mnuExamEnterResults, "SCEXER", , ControlText.MnuSExamEnterResultsTT, True))
      .Add(CommandIndexes.mnuExamAllocateCandidateNumbers.ToString, New MenuToolbarCommand("ExamAllocateCandidateNumbers", ControlText.MnuSExamAllocateCandidateNumbers, CommandIndexes.mnuExamAllocateCandidateNumbers, "SCEXAC", , ControlText.MnuSExamAllocateCandidateNumbersTT, True))
      .Add(CommandIndexes.mnuExamAllocateMarkers.ToString, New MenuToolbarCommand("ExamAllocateMarkers", ControlText.MnuSExamAllocateMarkers, CommandIndexes.mnuExamAllocateMarkers, "SCEXAM", , ControlText.MnuSExamAllocateMarkersTT, True))
      .Add(CommandIndexes.mnuExamApplyGrading.ToString, New MenuToolbarCommand("ExamApplyGrading", ControlText.MnuSExamApplyGrading, CommandIndexes.mnuExamApplyGrading, "SCEXAG", , ControlText.MnuSExamApplyGradingTT, True))
      .Add(CommandIndexes.mnuExamGenerateExemptionInvoices.ToString, New MenuToolbarCommand("ExamGenerateExemptionInvoices", ControlText.MnuSExamGenerateExemptionInvoices, CommandIndexes.mnuExamGenerateExemptionInvoices, "SCEXGE", , ControlText.MnuSExamGenerateExemptionInvoicesTT, True))
      .Add(CommandIndexes.mnuExamLoadCSVResults.ToString, New MenuToolbarCommand("ExamLoadCSVResults", ControlText.MnuSExamLoadCSVResults, CommandIndexes.mnuExamLoadCSVResults, "SCEXLC", , ControlText.MnuSExamLoadCSVResultsTT, True))
      .Add(CommandIndexes.mnuExamCancelProvisionalBookings.ToString, New MenuToolbarCommand("CancelProvisionalExamBookings", ControlText.MnuSExamCancelProvisionalBookings, CommandIndexes.mnuExamCancelProvisionalBookings, "SCEXCP", , ControlText.MnuSExamCancelProvisionalBookingsTT, True))
      .Add(CommandIndexes.mnuExamProcessCertificates.ToString, New MenuToolbarCommand("ExamProcessCertificates", ControlText.MnuSExamProcessCertificates, CommandIndexes.mnuExamProcessCertificates, "SCEXPC", , ControlText.MnuSExamProcessCertificatesTT, True))
      .Add(CommandIndexes.mnuExamGenerateCertificates.ToString, New MenuToolbarCommand("ExamGenerateCertificates", ControlText.MnuSExamGenerateCertificates, CommandIndexes.mnuExamGenerateCertificates, "SCEXGC", , ControlText.MnuSExamGenerateCertificatesTT, True))
      .Add(CommandIndexes.mnuExamSheduleCertificateReprints.ToString, New MenuToolbarCommand("ScheduleCertificateReprints", ControlText.MnuSExamScheduleCertificateReprints, CommandIndexes.mnuExamSheduleCertificateReprints, "SCEXGC", , ControlText.MnuSExamScheduleCertificateReprintsTT, True))

      .Add(CommandIndexes.mnuCpdApplyPoints.ToString, New MenuToolbarCommand("ApplyPoints", ControlText.MnuSApplyPoints, CommandIndexes.mnuCpdApplyPoints, "SCCPAP", , ControlText.MnuSApplyPoints, True))
      '-------------------------------------------------------------
      'ADMINISTRATION MENU
      '-------------------------------------------------------------

      .Add(CommandIndexes.mnuAdminAccessControl.ToString, New MenuToolbarCommand("AccessControl", ControlText.MnuAdminAccessControl, CommandIndexes.mnuAdminAccessControl, "SCAMAC", , ControlText.MnuAdminAccessControlTT, True))
      .Add(CommandIndexes.mnuAdminConfigurationMaintenance.ToString, New MenuToolbarCommand("ConfigurationMaintenance", ControlText.MnuAdminConfigurationMaintenance, CommandIndexes.mnuAdminConfigurationMaintenance, "SCAMCM", , ControlText.MnuAdminConfigurationMaintenanceTT, True))
      .Add(CommandIndexes.mnuAdminLicenceMaintenance.ToString, New MenuToolbarCommand("LicenceMaintenance", ControlText.MnuAdminLicenceMaintenance, CommandIndexes.mnuAdminLicenceMaintenance, "SCAMLM", , ControlText.MnuAdminLicenceMaintenanceTT, True))
      .Add(CommandIndexes.mnuAdminMaintenanceSetup.ToString, New MenuToolbarCommand("MaintenanceSetup", ControlText.MnuAdminMaintenanceSetup, CommandIndexes.mnuAdminMaintenanceSetup, "SCAMMS", , ControlText.MnuAdminMaintenanceSetupTT, True))
      .Add(CommandIndexes.mnuAdminMoveExternalDocuments.ToString, New MenuToolbarCommand("Move External Documents", ControlText.MnuAdminMoveExternalDocuments, CommandIndexes.mnuAdminMoveExternalDocuments, "SCAMEX", , ControlText.MnuAdminMoveExternalDocumentsTT, True))
      .Add(CommandIndexes.mnuAdminOwnershipMaintenance.ToString, New MenuToolbarCommand("OwnershipMaintenance", ControlText.MnuAdminOwnershipMaintenance, CommandIndexes.mnuAdminOwnershipMaintenance, "SCAMOM", , ControlText.MnuAdminOwnershipMaintenanceTT, True))
      .Add(CommandIndexes.mnuAdminReportMaintenance.ToString, New MenuToolbarCommand("ReportMaintenance", ControlText.MnuAdminReportMaintenance, CommandIndexes.mnuAdminReportMaintenance, "SCAMRM", , ControlText.MnuAdminReportMaintenanceTT, True))
      .Add(CommandIndexes.mnuAdminTraderApplicationMaintenance.ToString, New MenuToolbarCommand("TraderApplicationMaintenance", ControlText.MnuAdminTraderApplicationMaintenance, CommandIndexes.mnuAdminTraderApplicationMaintenance, "SCAMTM", , ControlText.MnuAdminTraderApplicationMaintenanceTT, True))
      .Add(CommandIndexes.mnuAdminDatabaseUpgrade.ToString, New MenuToolbarCommand("DatabaseUpgrade", ControlText.MnuAdminDatabaseUpgrade, CommandIndexes.mnuAdminDatabaseUpgrade, "SCAMDU", , ControlText.MnuAdminDatabaseUpgradeTT, True))
      .Add(CommandIndexes.mnuAdminDataImport.ToString, New MenuToolbarCommand("DataImport", ControlText.MnuAdminDataImport, CommandIndexes.mnuAdminDataImport, "SCAMDI", , ControlText.MnuAdminDataImportTT, True))
      .Add(CommandIndexes.mnuAdminUpdateCustomForms.ToString, New MenuToolbarCommand("UpdateCustomForms", ControlText.MnuAdminUpdateCustomForms, CommandIndexes.mnuAdminUpdateCustomForms, "SCAMUC", , ControlText.MnuAdminUpdateCustomFormsTT, True))
      .Add(CommandIndexes.mnuAdminUpdateGovernmentRegions.ToString, New MenuToolbarCommand("UpdateGovernmentRegions", ControlText.MnuAdminUpdateGovernmentRegions, CommandIndexes.mnuAdminUpdateGovernmentRegions, "SCAMUG", , ControlText.MnuAdminUpdateGovernmentRegionsTT, True))
      .Add(CommandIndexes.mnuAdminUpdateMailsortData.ToString, New MenuToolbarCommand("UpdateMailsortData", ControlText.MnuAdminUpdateMailsortData, CommandIndexes.mnuAdminUpdateMailsortData, "SCAMUD", , ControlText.MnuAdminUpdateMailsortDataTT, True))
      .Add(CommandIndexes.mnuAdminUpdatePaymentSchedule.ToString, New MenuToolbarCommand("UpdatePaymentSchedule", ControlText.MnuAdminUpdatePaymentSchedule, CommandIndexes.mnuAdminUpdatePaymentSchedule, "SCAMUP", , ControlText.MnuAdminUpdatePaymentScheduleTT, True))
      .Add(CommandIndexes.mnuAdminUpdateTraderApplications.ToString, New MenuToolbarCommand("UpdateTraderApplications", ControlText.MnuAdminUpdateTraderApplications, CommandIndexes.mnuAdminUpdateTraderApplications, "SCAMUT", , ControlText.MnuAdminUpdateTraderApplicationsTT, True))
      .Add(CommandIndexes.mnuAdminPostcodeUpdate.ToString, New MenuToolbarCommand("PostcodeUpdate", ControlText.MnuAdminPostcodeUpdate, CommandIndexes.mnuAdminPostcodeUpdate, "SCAMPU", , ControlText.MnuAdminPostcodeUpdateTT, True))
      .Add(CommandIndexes.mnuAdminDataUpdates.ToString, New MenuToolbarCommand("DataUpdates", ControlText.MnuAdminDataUpdates, CommandIndexes.mnuAdminDataUpdates, "SCAMDP", , ControlText.MnuAdminDataUpdatesTT, True))
      .Add(CommandIndexes.mnuAdminImportTraderApplication.ToString, New MenuToolbarCommand("ImportTraderApplication", ControlText.MnuAdminImportTraderApplication, CommandIndexes.mnuAdminImportTraderApplication, "SCAMIT", , ControlText.MnuAdminImportTraderApplicationTT, True))
      .Add(CommandIndexes.mnuAdminExportCustomForm.ToString, New MenuToolbarCommand("ExportCustomForm", ControlText.MnuAdminExportCustomForm, CommandIndexes.mnuAdminExportCustomForm, "SCAMEC", , ControlText.MnuAdminExportCustomFormTT, True))
      .Add(CommandIndexes.mnuAdminExportReport.ToString, New MenuToolbarCommand("ExportReport", ControlText.MnuAdminExportReport, CommandIndexes.mnuAdminExportReport, "SCAMER", , ControlText.MnuAdminExportReportTT, True))
      .Add(CommandIndexes.mnuAdminExportTraderApplication.ToString, New MenuToolbarCommand("ExportTraderApplication", ControlText.MnuAdminExportTraderApplication, CommandIndexes.mnuAdminExportTraderApplication, "SCAMET", , ControlText.MnuAdminExportTraderApplicationTT, True))
      .Add(CommandIndexes.mnuAdminConfigurationReport.ToString, New MenuToolbarCommand("ConfigurationReport", ControlText.MnuAdminConfigurationReport, CommandIndexes.mnuAdminConfigurationReport, "SCAMCR", , ControlText.MnuAdminConfigurationReportTT, True))
      .Add(CommandIndexes.mnuAdminCheckSetup.ToString, New MenuToolbarCommand("CheckSetup", ControlText.MnuAdminCheckSetup, CommandIndexes.mnuAdminCheckSetup, "SCAMCS", , ControlText.MnuAdminCheckSetupTT, True))
      .Add(CommandIndexes.mnuAdminCheckPaymentPlans.ToString, New MenuToolbarCommand("CheckPaymentPlans", ControlText.MnuAdminCheckPaymentPlans, CommandIndexes.mnuAdminCheckPaymentPlans, "SCAMCP", , ControlText.MnuAdminCheckPaymentPlansTT, True))
      .Add(CommandIndexes.mnuAdminRegenerateMessageQueue.ToString, New MenuToolbarCommand("RegenerateMessageQueue", ControlText.MnuAdminRegenerateMessageQueue, CommandIndexes.mnuAdminRegenerateMessageQueue, "SCAMRQ", , ControlText.MnuAdminRegenerateMessageQueueTT, True))
      .Add(CommandIndexes.mnuEventCancel.ToString, New MenuToolbarCommand("Cancel", ControlText.MnuEventCancel, CommandIndexes.mnuEventCancel, "SCEVCN", , ControlText.MnuEventCancel, True))
      .Add(CommandIndexes.mnuEventCancelProvisionalTransaction.ToString, New MenuToolbarCommand("CancelProvisionalBookings", ControlText.MnuEventCancelProvisionalBookings, CommandIndexes.mnuEventCancelProvisionalTransaction, "SCECPB", , ControlText.MnuEventCancelProvisionalBookings, True))

      .Add(CommandIndexes.cbiQueryByExample.ToString, New MenuToolbarCommand("QueryByExample", ControlText.MnuMQueryByExample, CommandIndexes.cbiQueryByExample, "SCQMBE", mvImageList16.Images(CommandIndexes.cbiQueryByExample), ControlText.MnuMQueryByExampleTT))
      .Add(CommandIndexes.cbiQueryByExampleContacts.ToString, New MenuToolbarCommand("QueryContact", ControlText.MnuMQueryContacts, CommandIndexes.cbiQueryByExampleContacts, "SCQMPF", mvImageList16.Images(CommandIndexes.cbiQueryByExampleContacts), "Query Contact"))
      .Add(CommandIndexes.cbiQueryByExampleContacts2.ToString, New MenuToolbarCommand("QueryContact2", "&Query Contact2...", CommandIndexes.cbiQueryByExampleContacts2, "SCQMP2", mvImageList16.Images(CommandIndexes.cbiQueryByExampleContacts2), "Query Contact2"))
      .Add(CommandIndexes.cbiQueryByExampleContacts3.ToString, New MenuToolbarCommand("QueryContact3", "&Query Contact3...", CommandIndexes.cbiQueryByExampleContacts3, "SCQMP3", mvImageList16.Images(CommandIndexes.cbiQueryByExampleContacts3), "Query Contact3"))
      .Add(CommandIndexes.cbiQueryByExampleContacts4.ToString, New MenuToolbarCommand("QueryContact4", "&Query Contact4...", CommandIndexes.cbiQueryByExampleContacts4, "SCQMP4", mvImageList16.Images(CommandIndexes.cbiQueryByExampleContacts4), "Query Contact4"))
      .Add(CommandIndexes.cbiQueryByExampleContacts5.ToString, New MenuToolbarCommand("QueryContact5", "&Query Contact5...", CommandIndexes.cbiQueryByExampleContacts5, "SCQMP5", mvImageList16.Images(CommandIndexes.cbiQueryByExampleContacts5), "Query Contact5"))
      .Add(CommandIndexes.cbiQueryByExampleOrganisations.ToString, New MenuToolbarCommand("QueryOrganisation", ControlText.MnuMQueryOrganisations, CommandIndexes.cbiQueryByExampleOrganisations, "SCQMOF", mvImageList16.Images(CommandIndexes.cbiQueryByExampleOrganisations), "Query Organisation"))
      .Add(CommandIndexes.cbiQueryByExampleOrganisations2.ToString, New MenuToolbarCommand("QueryOrganisation2", "&Query Organisation2...", CommandIndexes.cbiQueryByExampleOrganisations2, "SCQMO2", mvImageList16.Images(CommandIndexes.cbiQueryByExampleOrganisations2), "Query Organisation2"))
      .Add(CommandIndexes.cbiQueryByExampleOrganisations3.ToString, New MenuToolbarCommand("QueryOrganisation3", "&Query Organisation3...", CommandIndexes.cbiQueryByExampleOrganisations3, "SCQMO3", mvImageList16.Images(CommandIndexes.cbiQueryByExampleOrganisations3), "Query Organisation3"))
      .Add(CommandIndexes.cbiQueryByExampleOrganisations4.ToString, New MenuToolbarCommand("QueryOrganisation4", "&Query Organisation4...", CommandIndexes.cbiQueryByExampleOrganisations4, "SCQMO4", mvImageList16.Images(CommandIndexes.cbiQueryByExampleOrganisations4), "Query Organisation4"))
      .Add(CommandIndexes.cbiQueryByExampleOrganisations5.ToString, New MenuToolbarCommand("QueryOrganisation5", "&Query Organisation5...", CommandIndexes.cbiQueryByExampleOrganisations5, "SCQMO5", mvImageList16.Images(CommandIndexes.cbiQueryByExampleOrganisations5), "Query Organisation5"))
      .Add(CommandIndexes.cbiQueryByExampleEvents.ToString, New MenuToolbarCommand("QueryEvent", ControlText.MnuMQueryEvents, CommandIndexes.cbiQueryByExampleEvents, "SCQMEF", mvImageList16.Images(CommandIndexes.cbiQueryByExampleEvents), "Query Event"))
      .Add(CommandIndexes.cbiQueryByExampleEvents2.ToString, New MenuToolbarCommand("QueryEvent2", "&Query Event2...", CommandIndexes.cbiQueryByExampleEvents2, "SCQME2", mvImageList16.Images(CommandIndexes.cbiQueryByExampleEvents2), "Query Event2"))
      .Add(CommandIndexes.cbiQueryByExampleEvents3.ToString, New MenuToolbarCommand("QueryEvent3", "&Query Event3...", CommandIndexes.cbiQueryByExampleEvents3, "SCQME3", mvImageList16.Images(CommandIndexes.cbiQueryByExampleEvents3), "Query Event3"))
      .Add(CommandIndexes.cbiQueryByExampleEvents4.ToString, New MenuToolbarCommand("QueryEvent4", "&Query Event4...", CommandIndexes.cbiQueryByExampleEvents4, "SCQME4", mvImageList16.Images(CommandIndexes.cbiQueryByExampleEvents4), "Query Event4"))
      .Add(CommandIndexes.cbiQueryByExampleEvents5.ToString, New MenuToolbarCommand("QueryEvent5", "&Query Event5...", CommandIndexes.cbiQueryByExampleEvents5, "SCQME5", mvImageList16.Images(CommandIndexes.cbiQueryByExampleEvents5), "Query Event5"))

      .Add(CommandIndexes.mnuInternalCheckNonCoreTables.ToString, New MenuToolbarCommand("CheckNonCoreTables", "Check Non Core Tables", CommandIndexes.mnuInternalCheckNonCoreTables, "", , "Check Non Core Tables"))
      .Add(CommandIndexes.mnuInternalGenerateTableCreationFiles.ToString, New MenuToolbarCommand("GenerateTableCreationFiles", "Generate Table Creation Files", CommandIndexes.mnuInternalGenerateTableCreationFiles, "", , "Generate Table Creation Files"))
      .Add(CommandIndexes.mnuInternalGetReportData.ToString, New MenuToolbarCommand("GetReportData", "Get Report Data", CommandIndexes.mnuInternalGetReportData, "", , "Get Report Data"))
      .Add(CommandIndexes.mnuInternalGetConfigNameData.ToString, New MenuToolbarCommand("GetConfigNameData", "Get Config Name Data", CommandIndexes.mnuInternalGetConfigNameData, "", , "Get Config Name Data"))

    End With

    mvMenuItems(CommandIndexes.cbiDashboard.ToString).DefaultNotVisible = True
    mvMenuItems(CommandIndexes.cbiHeaderPanel.ToString).DefaultNotVisible = True
    mvMenuItems(CommandIndexes.cbiSelectionPanel.ToString).DefaultNotVisible = True
    mvMenuItems(CommandIndexes.cbiSearchData.ToString).DefaultNotVisible = True

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems, True)

    'This menu item is currently not implemented so always hide even if user made it visible
    mvMenuItems(CommandIndexes.mnuAdminPostcodeUpdate.ToString).HideItem = True

    mvMenuItems(CommandIndexes.cbiLogin.ToString).AddToMenu(FileToolStripMenuItem)
    Dim vFileNew As New ToolStripMenuItem(ControlText.MnuMNew)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtContactGroup, vFileNew, GroupMenuItemTypes.gmiNewContact)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtOrganisationGroup, vFileNew, GroupMenuItemTypes.gmiNewOrganisation)
    mvMenuItems(CommandIndexes.cbiNewDocument.ToString).AddToMenu(vFileNew)
    mvMenuItems(CommandIndexes.cbiNewAction.ToString).AddToMenu(vFileNew)
    mvMenuItems(CommandIndexes.cbiNewActionTemplate.ToString).AddToMenu(vFileNew)
    mvMenuItems(CommandIndexes.cbiNewTelephoneCall.ToString).AddToMenu(vFileNew)
    mvMenuItems(CommandIndexes.cbiNewSelectionSet.ToString).AddToMenu(vFileNew)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtEventGroup, vFileNew, GroupMenuItemTypes.gmiNewEvent)
    If vFileNew.DropDownItems.Count > 0 Then FileToolStripMenuItem.DropDownItems.Add(vFileNew)
    mvMenuItems(CommandIndexes.cbiPreferences.ToString).AddToMenu(FileToolStripMenuItem)
    If Debugger.IsAttached Then mvMenuItems(CommandIndexes.cbiLogWEBServices.ToString).AddToMenu(FileToolStripMenuItem)
    FileToolStripMenuItem.DropDownItems.Add(New ToolStripSeparator)
    mvMenuItems(CommandIndexes.cbiPageSetup.ToString).AddToMenu(FileToolStripMenuItem)
    FileToolStripMenuItem.DropDownItems.Add(New ToolStripSeparator)
    mvMenuItems(CommandIndexes.cbiExit.ToString).AddToMenu(FileToolStripMenuItem)

    mvMenuItems(CommandIndexes.cbiNextRecord.ToString).AddToMenu(ViewToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiPreviousRecord.ToString).AddToMenu(ViewToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiTableMaintenance.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiJobSchedule.ToString).AddToMenu(ToolsToolStripMenuItem)

    mvMenuItems(CommandIndexes.cbiToolbar.ToString).AddToMenu(ViewToolStripMenuItem, True)
    mvMenuItems(CommandIndexes.cbiNavigationPanel.ToString).AddToMenu(ViewToolStripMenuItem, True)
    mvMenuItems(CommandIndexes.cbiStatusBar.ToString).AddToMenu(ViewToolStripMenuItem, True)
    mvMenuItems(CommandIndexes.cbiHeaderPanel.ToString).AddToMenu(ViewToolStripMenuItem, True)
    mvMenuItems(CommandIndexes.cbiSelectionPanel.ToString).AddToMenu(ViewToolStripMenuItem, True)
    mvMenuItems(CommandIndexes.cbiDashboard.ToString).AddToMenu(ViewToolStripMenuItem, False)
    Dim vViewMy As New ToolStripMenuItem(ControlText.MnuMMy)
    mvMenuItems(CommandIndexes.cbiMyDetails.ToString).AddToMenu(vViewMy)
    mvMenuItems(CommandIndexes.cbiMyOrganisation.ToString).AddToMenu(vViewMy)
    mvMenuItems(CommandIndexes.cbiMyActions.ToString).AddToMenu(vViewMy)
    mvMenuItems(CommandIndexes.cbiMyDocuments.ToString).AddToMenu(vViewMy)
    mvMenuItems(CommandIndexes.cbiMyInBox.ToString).AddToMenu(vViewMy)
    mvMenuItems(CommandIndexes.cbiMyJournal.ToString).AddToMenu(vViewMy)
    If vViewMy.DropDownItems.Count > 0 Then ViewToolStripMenuItem.DropDownItems.Add(vViewMy)
    mvMenuItems(CommandIndexes.cbiRefresh.ToString).AddToMenu(ViewToolStripMenuItem)

    mvMenuItems(CommandIndexes.cbiQueryByExample.ToString).AddToMenu(QueryToolStripMenuItem)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtContactGroup, QueryToolStripMenuItem, GroupMenuItemTypes.gmiQueryContact)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtOrganisationGroup, QueryToolStripMenuItem, GroupMenuItemTypes.gmiQueryOrganisation)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtEventGroup, QueryToolStripMenuItem, GroupMenuItemTypes.gmiQueryEvent)

    mvMenuItems(CommandIndexes.cbiSearchData.ToString).AddToMenu(FindToolStripMenuItem)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtContactGroup, FindToolStripMenuItem, GroupMenuItemTypes.gmiFindContact)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtOrganisationGroup, FindToolStripMenuItem, GroupMenuItemTypes.gmiFindOrganisation)
    mvMenuItems(CommandIndexes.cbiDocumentFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiMeeting.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiActionFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiSelectionSetFinder.ToString).AddToMenu(FindToolStripMenuItem)
    AddGroupMenuItems(EntityGroup.EntityGroupTypes.egtEventGroup, FindToolStripMenuItem, GroupMenuItemTypes.gmiFindEvent)
    mvMenuItems(CommandIndexes.cbiMemberFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiPayPlanFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiCovenantFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiXactionFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiStandingOrderFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiDirectDebitFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiCCCAFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiGADFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiInvoiceFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiLegacyFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiGAYEFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiPostTaxPGFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiPurchaseOrderFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiProductFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiCampaignFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiStandardDocument.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiFundraisingPaymentScheduleFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiServiceProductFinder.ToString).AddToMenu(FindToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiFundraisingRequestFinder.ToString).AddToMenu(FindToolStripMenuItem)

    mvMenuItems(CommandIndexes.cbiListManager.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiRunMailing.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiRunReport.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiDocumentDistributor.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiSendEmail.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiPostcodeProximity.ToString).AddToMenu(ToolsToolStripMenuItem)

    If mvMenuItems(CommandIndexes.cbiExplore.ToString).HideItem = False Then
      Dim vExplore As New ToolStripMenuItem(ControlText.MnuMExplore)
      Dim vParams As New ParameterList(True)
      vParams("Tools") = "Y"
      Dim vExplorerTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExplorerLinks, vParams)
      If Not vExplorerTable Is Nothing Then
        For Each vRow As DataRow In vExplorerTable.Rows
          Dim vExploreItem As ToolStripItem = vExplore.DropDownItems.Add(vRow("ExplorerLinkDesc").ToString)
          vExploreItem.Tag = vRow("ExplorerLinkUrl").ToString & "," & vRow("ShowToolbar").ToString
          vExplore.DropDownItems.Add(vExploreItem)
          AddHandler vExploreItem.Click, AddressOf ProcessExplorerItemFromMenu
        Next
      End If
      If vExplore.DropDownItems.Count > 0 Then ToolsToolStripMenuItem.DropDownItems.Add(vExplore)
    End If
    mvMenuItems(CommandIndexes.cbiCustomise.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiCloseOpenBatch.ToString).AddToMenu(ToolsToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiCopyEventPricingMatrix.ToString).AddToMenu(ToolsToolStripMenuItem)
    If System.Diagnostics.Debugger.IsAttached Then
      mvMenuItems(CommandIndexes.cbiRunTest.ToString).AddToMenu(ToolsToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiClearCache.ToString).AddToMenu(ToolsToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiAllowCaching.ToString).AddToMenu(ToolsToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiAllowCaching.ToString).CheckToolStripItem(mvMenuStrip, Nothing, AppValues.AllowCaching)
    Else
      'SystemToolStripMenuItem.Visible = False
    End If

    Dim vBankItems() As CommandIndexes = {CommandIndexes.mnuFinBanksLoad}
    AddSubMenuItems(vBankItems, ControlText.MnuSBanks, SystemToolStripMenuItem)

    Dim vBatchItems() As CommandIndexes
    Dim vBatchItems1() As CommandIndexes = {
        CommandIndexes.mnuFinBViewBatchDetail,
        CommandIndexes.mnuFinBProcessBatches,
        CommandIndexes.mnuFinBOutstandingBatchesReport,
        CommandIndexes.mnuFinBChequeList,
        CommandIndexes.mnuFinBPayingInSlips,
        CommandIndexes.mnuFinBCashBookBatch,
        CommandIndexes.mnuFinBUpdateBatch,
        CommandIndexes.mnuFinBCreateJournal,
        CommandIndexes.mnuFinBSummaryReport,
        CommandIndexes.mnuFinBDetailReport,
        CommandIndexes.mnuFinBPurgeOldBatches,
        CommandIndexes.mnuFinBPurgePrizeDrawBatches}
    vBatchItems = vBatchItems1

    AddSubMenuItems(vBatchItems, ControlText.MnuSBatchManagement, SystemToolStripMenuItem)

    Dim vCAFItems() As CommandIndexes = {
      CommandIndexes.mnuFinCAFExpectedPayments,
      CommandIndexes.mnuFinCAFProvisionalBatchClaim,
      CommandIndexes.mnuFinCAFCreateCardSalesReport,
      CommandIndexes.mnuFinCAFLoadPaymentData,
      CommandIndexes.mnuFinCAFReconcilePaymentData}
    AddSubMenuItems(vCAFItems, ControlText.MnuSCAF, SystemToolStripMenuItem)

    Dim vCPDItems() As CommandIndexes = {CommandIndexes.mnuCpdApplyPoints}
    AddSubMenuItems(vCPDItems, ControlText.MnuSCpd, SystemToolStripMenuItem)

    Dim vCCItems() As CommandIndexes = {
      CommandIndexes.mnuFinCCBatches,
      CommandIndexes.mnuFinCCCreateFile,
      CommandIndexes.mnuFinCCCreateReport,
      CommandIndexes.mnuFinCCCreateCardSalesFile,
      CommandIndexes.mnuFinCCCreateCardSalesReport,
      CommandIndexes.mnuFinCCAuthorisationsReport}
    AddSubMenuItems(vCCItems, ControlText.MnuSCreditCards, SystemToolStripMenuItem)

    Dim vCSItems() As CommandIndexes = {
      CommandIndexes.mnuFinCSTransferInvoices,
      CommandIndexes.mnuFinCSTransferCustomers,
      CommandIndexes.mnuFinCSStatementGeneration}
    AddSubMenuItems(vCSItems, ControlText.MnuSCreditSales, SystemToolStripMenuItem)

    'Fast Data Entry
    Dim vFDEItems() As CommandIndexes = {CommandIndexes.mnuFinFastDataEntryMaint}
    AddSubMenuItems(vFDEItems, ControlText.MnuSFDE, SystemToolStripMenuItem)

    Dim vDeDupItems() As CommandIndexes = {
        CommandIndexes.mnuDeDuplicationContactMerge,
        CommandIndexes.mnuDeDuplicationAddressMerge,
        CommandIndexes.mnuDeDuplicationOrganisationMerge,
        CommandIndexes.mnuDeDuplicationAmalgamateOrganisations,
        CommandIndexes.mnuDeDuplicationContactDeDuplication,
        CommandIndexes.mnuDeDuplicationBulkAddressMerge,
        CommandIndexes.mnuDeDuplicationBulkContactMerge,
        CommandIndexes.mnuDeDuplicationBulkOrganisationMerge,
        CommandIndexes.mnuDeDuplicateProcessDuplicateContacts}
    AddSubMenuItems(vDeDupItems, ControlText.MnuSDeDuplication, SystemToolStripMenuItem)

    Dim vDDItems() As CommandIndexes = {
      CommandIndexes.mnuFinDDMandateFile,
      CommandIndexes.mnuFinDDBatches,
      CommandIndexes.mnuFinDDClaimFile,
      CommandIndexes.mnuFinDDDirectCreditFile,
      CommandIndexes.mnuFinDDUploadBacsMessagingData,
      CommandIndexes.mnuFinDDBACSRejections,
      CommandIndexes.mnuFinDDConvertManualDirectDebits}
    AddSubMenuItems(vDDItems, ControlText.MnuSDirectDebits, SystemToolStripMenuItem)

    Dim vDistributionBoxes As New ToolStripMenuItem(ControlText.MnuSDistributionBoxes)
    Dim vDBReportItems() As CommandIndexes = {
      CommandIndexes.mnuFinDBRepOpenBoxes,
      CommandIndexes.mnuFinDBRepUnAllocatedDonations,
      CommandIndexes.mnuFinDBRepAllocatedDonations,
      CommandIndexes.mnuFinDBRepDonorDetails,
      CommandIndexes.mnuFinDBRepClosedByLocation,
      CommandIndexes.mnuFinDBRepRollOfHonour}
    AddSubMenuItems(vDBReportItems, ControlText.MnuSReports, vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBCreateUnallocatedBoxes.ToString).AddToMenu(vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBPrintThankYouLetters.ToString).AddToMenu(vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBPrintAdviceNotes.ToString).AddToMenu(vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBPrintPackingSlips.ToString).AddToMenu(vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBPrintBoxLabels.ToString).AddToMenu(vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBSetShippingInformation.ToString).AddToMenu(vDistributionBoxes)
    mvMenuItems(CommandIndexes.mnuFinDBSetArrivalInformation.ToString).AddToMenu(vDistributionBoxes)
    If vDistributionBoxes.DropDownItems.Count > 0 Then SystemToolStripMenuItem.DropDownItems.Add(vDistributionBoxes)

    Dim vDutchItems() As CommandIndexes = {
         CommandIndexes.mnuDutchLoadPayments,
         CommandIndexes.mnuDutchProcessPayments}
    AddSubMenuItems(vDutchItems, ControlText.MnuSDutchPaymentProcessing, SystemToolStripMenuItem)

    Dim vCMItems() As CommandIndexes = {
         CommandIndexes.mnuCMProcessMetaData,
         CommandIndexes.mnuCMProcessEventData,
         CommandIndexes.mnuCMProcessTotals,
         CommandIndexes.mnuCMUpdateBulkMailer}
    AddSubMenuItems(vCMItems, ControlText.MnuSCMCheetahMail, SystemToolStripMenuItem)

    Dim vEventItems() As CommandIndexes = {CommandIndexes.mnuEventBlockBooking, CommandIndexes.mnuEventCancel, CommandIndexes.mnuEventCancelProvisionalTransaction}
    AddSubMenuItems(vEventItems, ControlText.MnuSEvents, SystemToolStripMenuItem)

    Dim vExamItems() As CommandIndexes = {
      CommandIndexes.mnuExamsMaintenance,
      CommandIndexes.mnuExamEnterResults,
      CommandIndexes.mnuExamAllocateCandidateNumbers,
      CommandIndexes.mnuExamAllocateMarkers,
      CommandIndexes.mnuExamApplyGrading,
      CommandIndexes.mnuExamGenerateExemptionInvoices,
      CommandIndexes.mnuExamLoadCSVResults,
      CommandIndexes.mnuExamCancelProvisionalBookings,
      CommandIndexes.mnuExamProcessCertificates,
      CommandIndexes.mnuExamGenerateCertificates,
      CommandIndexes.mnuExamSheduleCertificateReprints}
    AddSubMenuItems(vExamItems, ControlText.MnuSExams, SystemToolStripMenuItem)

    Dim vGiftAid As New ToolStripMenuItem(ControlText.MnuSGiftAid)
    Dim vGADItems() As CommandIndexes = {
      CommandIndexes.mnuFinGADConfirmation,
      CommandIndexes.mnuFinGADPotentialClaim,
      CommandIndexes.mnuFinGADTaxClaim,
      CommandIndexes.mnuFinGADGiftDetails,
      CommandIndexes.mnuFinGADGiftAnalysis,
      CommandIndexes.mnuFinBulkGiftAidUpdate}
    AddSubMenuItems(vGADItems, ControlText.MnuSDeclarationsCovenants, vGiftAid)
    Dim vGASItems() As CommandIndexes = {
      CommandIndexes.mnuFinGASPotentialClaim,
      CommandIndexes.mnuFinGASTaxClaim,
      CommandIndexes.mnuFinGASClaimDetails}
    AddSubMenuItems(vGASItems, ControlText.MnuSSponsorship, vGiftAid)
    Dim vGAIItems() As CommandIndexes = {
          CommandIndexes.mnuFinGAIPotentialClaim,
          CommandIndexes.mnuFinGAITaxClaim,
          CommandIndexes.mnuFinGAIClaimDetails}
    AddSubMenuItems(vGAIItems, ControlText.MnuSIrish, vGiftAid)
    If vGiftAid.DropDownItems.Count > 0 Then SystemToolStripMenuItem.DropDownItems.Add(vGiftAid)

    Dim vIncItems() As CommandIndexes = {CommandIndexes.mnuFinancialIncentives}
    AddSubMenuItems(vIncItems, ControlText.MnuSIncentives, SystemToolStripMenuItem)

    Dim vMailingDocuments As New ToolStripMenuItem(ControlText.MnuSMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingDocsProduce.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingDocsFind.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingTYL.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingFinder.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuEMailProcessor.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingListAllContacts.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingDirectDebit.ToString).AddToMenu(vMailingDocuments)
    Dim vMailingEvents() As CommandIndexes = {
     CommandIndexes.mnuMailingEventsBookings,
     CommandIndexes.mnuMailingEventsDelegates,
     CommandIndexes.mnuMailingEventsPersonnel,
     CommandIndexes.mnuMailingEventsSponsors}
    AddSubMenuItems(vMailingEvents, ControlText.MnuSMailingEvents, vMailingDocuments)
    Dim vMailingExamss() As CommandIndexes = {
     CommandIndexes.mnuMailingExamsBookings,
     CommandIndexes.mnuMailingExamsCandidates}
    AddSubMenuItems(vMailingExamss, ControlText.MnuSMailingExams, vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingIrishGiftAid.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingMembers.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingPayers.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingPreTextPayrollGivingPledges.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingSelectionManager.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingSelectionTester.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingStandingOrders.ToString).AddToMenu(vMailingDocuments)
    mvMenuItems(CommandIndexes.mnuMailingSubscriptions.ToString).AddToMenu(vMailingDocuments)

    If vMailingDocuments.DropDownItems.Count > 0 Then SystemToolStripMenuItem.DropDownItems.Add(vMailingDocuments)

    Dim vMKTItems() As CommandIndexes = {CommandIndexes.mnuMktGenerateData}
    AddSubMenuItems(vMKTItems, ControlText.MnuSMarketing, SystemToolStripMenuItem)

    Dim vMembership As New ToolStripMenuItem(ControlText.MnusMembership)
    mvMenuItems(CommandIndexes.mnuMemFutureChanges.ToString).AddToMenu(vMembership)
    mvMenuItems(CommandIndexes.mnuMemCards.ToString).AddToMenu(vMembership)
    mvMenuItems(CommandIndexes.mnuMemSuspension.ToString).AddToMenu(vMembership)
    mvMenuItems(CommandIndexes.mnuMemFulfilment.ToString).AddToMenu(vMembership)
    Dim vMemReports() As CommandIndexes = {
      CommandIndexes.mnuMemAssumedVotingRights,
      CommandIndexes.mnuMemBallotPaperProduction,
      CommandIndexes.mnuMemBranchDonations,
      CommandIndexes.mnuMemBranchIncome,
      CommandIndexes.mnuMemJuniorAnalysis}
    AddSubMenuItems(vMemReports, ControlText.MnuSMemReports, vMembership)
    Dim vMemStats() As CommandIndexes = {
      CommandIndexes.mnuMemGenerateStatistics,
      CommandIndexes.mnuMemStatisticsDetailed,
      CommandIndexes.mnuMemStatisticsSummary}
    AddSubMenuItems(vMemStats, ControlText.MnuSMemStatistics, vMembership)
    If vMembership.DropDownItems.Count > 0 Then SystemToolStripMenuItem.DropDownItems.Add(vMembership)

    Dim vNomItems() As CommandIndexes = {
      CommandIndexes.mnuFinNomSummaryReport,
      CommandIndexes.mnuFinNomDetailReport}
    AddSubMenuItems(vNomItems, ControlText.MnuSNominalCodes, SystemToolStripMenuItem)

    Dim vPISItems() As CommandIndexes = {
      CommandIndexes.mnuFinPISLoadStatement,
      CommandIndexes.mnuFinPISReconciliation}
    AddSubMenuItems(vPISItems, ControlText.MnuSPayingInSlips, SystemToolStripMenuItem)

    Dim vPPItems() As CommandIndexes = {
      CommandIndexes.mnuFinPPRenewals,
      CommandIndexes.mnuFinPPRemoveArrears,
      CommandIndexes.mnuFinPPExpiry,
      CommandIndexes.mnuFinPPNonMemberFulfilment,
      CommandIndexes.mnuFinPPUpdateProducts,
      CommandIndexes.mnuFinPPApplySurcharges,
      CommandIndexes.mnuFinPPReCalcLoanInterest,
      CommandIndexes.mnuFinPPUpdateLoanInterestRates,
      CommandIndexes.mnuFinPPTransferPaymentPlanChanges}
    AddSubMenuItems(vPPItems, ControlText.MnuSPaymentPlans, SystemToolStripMenuItem)

    Dim vPGItems() As CommandIndexes = {
      CommandIndexes.mnuFinGAYELoadPayments,
      CommandIndexes.mnuFinGAYEReconciliation,
      CommandIndexes.mnuFinGAYEBulkCancellation,
      CommandIndexes.mnuFinGAYEPostTaxPGReconciliation}
    AddSubMenuItems(vPGItems, ControlText.MnuSPayrollGiving, SystemToolStripMenuItem)

    Dim vPOItems() As CommandIndexes = {
      CommandIndexes.mnuFinPOAuthorisePayments,
      CommandIndexes.mnuFinPOAutoGenerate,
      CommandIndexes.mnuFinChequeProduction,
      CommandIndexes.mnuFinPOProcessPayments,
      CommandIndexes.mnuFinPOPrint,
      CommandIndexes.mnuFinPOTransSuppliers,
      CommandIndexes.mnuFinPOTransPayments}
    AddSubMenuItems(vPOItems, ControlText.MnuSPurchaseOrders, SystemToolStripMenuItem)

    Dim vProductItems() As CommandIndexes = {
      CommandIndexes.mnuFinPRPriceChange,
      CommandIndexes.mnuFinPRPurchasedProductReport}
    AddSubMenuItems(vProductItems, ControlText.MnuSProducts, SystemToolStripMenuItem)

    Dim vSOItems() As CommandIndexes = {
      CommandIndexes.mnuFinSOLoadStatement,
      CommandIndexes.mnuFinSOReconciliation,
      CommandIndexes.mnuFinSOCancellation,
      CommandIndexes.mnuFinSOManualRec,
      CommandIndexes.mnuFinSOReconReport,
      CommandIndexes.mnuFinSOBankTransactionsReport}
    AddSubMenuItems(vSOItems, ControlText.MnuSStandingOrders, SystemToolStripMenuItem)
    'CommandIndexes.mnuFinSTExport, _
    Dim vStockItems() As CommandIndexes = {
      CommandIndexes.mnuFinSTPickingLists,
      CommandIndexes.mnuFinSTConfirmAllocation,
      CommandIndexes.mnuFinSTAllocateToBO,
      CommandIndexes.mnuFinSTDespatchNotes,
      CommandIndexes.mnuFinSTDespatchTracking,
      CommandIndexes.mnuFinSTBackOrdersReport,
      CommandIndexes.mnuFinSTSalesAnalysis,
      CommandIndexes.mnuFinSTSalesAnalysisDetailed,
      CommandIndexes.mnuFinSTSalesAnalysisSummary,
      CommandIndexes.mnuFinSTMovement,
      CommandIndexes.mnuFinSTTransferStockToPack,
      CommandIndexes.mnuFinSTPurgeBackOrders,
      CommandIndexes.mnuFinSTPurgePickingAndDespatch,
      CommandIndexes.mnuFinSTPLAwaitingConfirm,
      CommandIndexes.mnuFinSTValuationReport}
    AddSubMenuItems(vStockItems, ControlText.MnuSStock, SystemToolStripMenuItem)
    'CommandIndexes.mnuAdminDeleteDormantContacts, _
    Dim vUpdateItems() As CommandIndexes = {
      CommandIndexes.mnuAdminAmendmentHistory,
      CommandIndexes.mnuAdminProcessAddressChanges,
      CommandIndexes.mnuAdminSetPostDatedContacts,
      CommandIndexes.mnuAdminUpdateActionStatuses,
      CommandIndexes.mnuAdminUpdateMailsort,
      CommandIndexes.mnuAdminUpdatePrincipalUser,
      CommandIndexes.mnuAdminUpdateRegionalData,
      CommandIndexes.mnuAdminUpdateSearchNames,
      CommandIndexes.mnuAdminPostcodeValidation,
      CommandIndexes.mnuAdminPurgeStickyNotes}
    AddSubMenuItems(vUpdateItems, ControlText.MnuSUpdates, SystemToolStripMenuItem)

    AdminToolStripMenuItem.Visible = False

    For vIndex As CommandIndexes = CommandIndexes.mnuAdminAccessControl To CommandIndexes.mnuAdminUpdateTraderApplications
      If mvMenuItems(vIndex.ToString).HideItem = False Then
        mvMenuItems(vIndex.ToString).AddToMenu(AdminToolStripMenuItem)
        AdminToolStripMenuItem.Visible = True
      End If
    Next

    If System.Diagnostics.Debugger.IsAttached Then
      Dim vInternalItems() As CommandIndexes = {
        CommandIndexes.mnuInternalCheckNonCoreTables,
        CommandIndexes.mnuInternalGenerateTableCreationFiles,
        CommandIndexes.mnuInternalGetReportData,
        CommandIndexes.mnuInternalGetConfigNameData}
      AddSubMenuItems(vInternalItems, "Internal", AdminToolStripMenuItem)
    End If

    ' This following block handles the system menu visibility
    ' Add to case statement as we add new features
    For vIndex As Integer = 0 To SystemToolStripMenuItem.DropDownItems.Count - 1
      SystemToolStripMenuItem.DropDownItems(vIndex).Visible = True
    Next
    If mvMenuStrip.MdiWindowListItem Is WindowToolStripMenuItem Then
      mvMenuItems(CommandIndexes.cbiCascade.ToString).AddToMenu(WindowToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiTileHorizontally.ToString).AddToMenu(WindowToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiTileVertically.ToString).AddToMenu(WindowToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiArrangeIcons.ToString).AddToMenu(WindowToolStripMenuItem)
      mvMenuItems(CommandIndexes.cbiCloseAll.ToString).AddToMenu(WindowToolStripMenuItem)
    Else
      mvMenuItems(CommandIndexes.cbiCascade.ToString).HideItem = True
      mvMenuItems(CommandIndexes.cbiTileHorizontally.ToString).HideItem = True
      mvMenuItems(CommandIndexes.cbiTileVertically.ToString).HideItem = True
      mvMenuItems(CommandIndexes.cbiArrangeIcons.ToString).HideItem = True
      mvMenuItems(CommandIndexes.cbiCloseAll.ToString).HideItem = True
    End If

    mvMenuItems(CommandIndexes.cbiHelp.ToString).AddToMenu(HelpToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiReleaseNotes.ToString).AddToMenu(HelpToolStripMenuItem)
    HelpToolStripMenuItem.DropDownItems.Add(New ToolStripSeparator)
    mvMenuItems(CommandIndexes.cbiKnowledgebase.ToString).AddToMenu(HelpToolStripMenuItem)
    mvMenuItems(CommandIndexes.cbiSupportForum.ToString).AddToMenu(HelpToolStripMenuItem)
    HelpToolStripMenuItem.DropDownItems.Add(New ToolStripSeparator)
    mvMenuItems(CommandIndexes.cbiAbout.ToString).AddToMenu(HelpToolStripMenuItem)

    If DataHelper.UserInfo.AccessLevel > UserInfo.UserAccessLevel.ualReadOnly Then
      'Fast Data Entry
      Dim vGotFDE As Boolean
      Dim vFDETable As DataTable = DataHelper.GetFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePages)
      If vFDETable IsNot Nothing Then
        For Each vRow As DataRow In vFDETable.Rows
          Dim vFDEItem As ToolStripItem = mvApplicationsMenu.DropDownItems.Add(vRow("FdePageName").ToString)
          vFDEItem.Tag = vRow("FdePageNumber").ToString
          AddHandler vFDEItem.Click, AddressOf ProcessFastDataEntryFromMenu
          vGotFDE = True
          If FormView = FormViews.Modern Then
            Dim vMenuCommand As New MenuToolbarCommand(vRow("FdePageNumber").ToString, vRow("FdePageName").ToString, IntegerValue(vRow("FdePageNumber").ToString))
            vMenuCommand.ExplorerMenuAttribute = New ExplorerMenuAttribute(ExplorerMenuSection.Trader, ExplorerMenuCategory.FastDataEntry)
            vMenuCommand.OnClick = AddressOf ProcessFastDataEntryFromCommand
            vMenuCommand.SerialisationFormat = "FDE{0}"
            mvMenuItems.Add(String.Format("FDE{0}", vRow("FdePageNumber").ToString), vMenuCommand)
          End If
        Next
      End If
      'Trader
      Dim vTraderAppTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtTraderApplications)
      If Not vTraderAppTable Is Nothing Then
        Dim vTraderAppList As New SortedList
        If vGotFDE = True AndAlso vTraderAppTable.Rows.Count > 0 Then
          Dim vItem As New ToolStripSeparator()
          mvApplicationsMenu.DropDownItems.Add(vItem)
        End If
        Dim vTraderConfig As String
        For Each vRow As DataRow In vTraderAppTable.Rows
          vTraderConfig = Mid(vRow("ConfigName").ToString, 21, 3)
          vTraderAppList.Add(IntegerValue(vTraderConfig), vRow("TraderApplication").ToString & "-" & vRow("TraderApplicationDesc").ToString)
        Next
        Dim vTraderAppNumber As String = ""
        Dim vTraderAppDesc As String = ""
        Dim vMenuItems As New List(Of ToolStripItem)
        For Each vItem As DictionaryEntry In vTraderAppList
          vTraderAppNumber = Mid(vItem.Value.ToString, 1, InStr(vItem.Value.ToString, "-") - 1)
          vTraderAppDesc = Mid(vItem.Value.ToString, InStr(vItem.Value.ToString, "-") + 1, vItem.Value.ToString.Length)
          'Dim vAppItem As ToolStripItem = mvApplicationsMenu.DropDownItems.Add(vTraderAppDesc)
          Dim vAppItem As New ToolStripMenuItem(vTraderAppDesc)
          vMenuItems.Add(vAppItem)
          vAppItem.Tag = vTraderAppNumber
          AddHandler vAppItem.Click, AddressOf LaunchTraderApplicationFromMenu
          If FormView = FormViews.Modern Then
            Dim vMenuCommand As New MenuToolbarCommand(vTraderAppNumber.ToString, vTraderAppDesc, IntegerValue(vTraderAppNumber))
            vMenuCommand.ExplorerMenuAttribute = New ExplorerMenuAttribute(ExplorerMenuSection.Trader, ExplorerMenuCategory.Trader)
            vMenuCommand.OnClick = AddressOf LaunchTraderApplicationFromCommand
            vMenuCommand.SerialisationFormat = "TRA{0}"
            mvMenuItems.Add(vItem.Value.ToString, vMenuCommand)
          End If
        Next
        mvApplicationsMenu.DropDownItems.AddRange(vMenuItems.ToArray)
      End If
      mvApplicationsMenu.Enabled = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_trader, True)
      'explorer links
      Dim vParams As New ParameterList(True)
      vParams("Tools") = "N"
      Dim vAppExplorerLinks As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExplorerLinks, vParams)
      If Not vAppExplorerLinks Is Nothing AndAlso vAppExplorerLinks.Columns.Contains("ExplorerLocation") Then
        If vAppExplorerLinks.Rows.Count > 0 Then
          'Add all the items in the middle first
          For Each vRow As DataRow In vAppExplorerLinks.Rows
            Dim vIndex As Integer
            Select Case vRow("ExplorerLocation").ToString
              Case "M"
                If vFDETable IsNot Nothing Then
                  vIndex = vFDETable.Rows.Count
                Else
                  vIndex = 0
                End If
            End Select
            If vRow("ExplorerLocation").ToString = "M" Then
              Dim vExploreItem As ToolStripItem = mvApplicationsMenu.DropDownItems.Add(vRow("ExplorerLinkDesc").ToString)
              Dim vTag As String = vRow("ExplorerLinkUrl").ToString & "," & vRow("ShowToolbar").ToString
              vExploreItem.Tag = vTag
              mvApplicationsMenu.DropDownItems.Insert(vIndex, vExploreItem)
              AddHandler vExploreItem.Click, AddressOf ProcessExplorerItemFromMenu
              If FormView = FormViews.Modern Then
                Dim vMenuCommand As New MenuToolbarCommand(vTag, vRow("ExplorerLinkDesc").ToString, 0)
                vMenuCommand.ExplorerMenuAttribute = New ExplorerMenuAttribute(ExplorerMenuSection.Trader, ExplorerMenuCategory.ExplorerMenuLinks)
                vMenuCommand.OnClick = AddressOf ProcessExplorerItemFromCommand
                vMenuCommand.SerialisationFormat = "EXP{0}"
                mvMenuItems.Add(vTag, vMenuCommand)
              End If
            End If
          Next
          'loop through and add all the top and bottom item
          For Each vRow As DataRow In vAppExplorerLinks.Rows
            Dim vIndex As Integer
            Select Case vRow("ExplorerLocation").ToString
              Case "T"
                vIndex = 0
              Case "B"
                vIndex = mvApplicationsMenu.DropDownItems.Count
            End Select
            If vRow("ExplorerLocation").ToString = "T" OrElse vRow("ExplorerLocation").ToString = "B" Then
              Dim vExploreItem As ToolStripItem = mvApplicationsMenu.DropDownItems.Add(vRow("ExplorerLinkDesc").ToString)
              Dim vTag As String = vRow("ExplorerLinkUrl").ToString & "," & vRow("ShowToolbar").ToString
              vExploreItem.Tag = vTag
              mvApplicationsMenu.DropDownItems.Insert(vIndex, vExploreItem)
              AddHandler vExploreItem.Click, AddressOf ProcessExplorerItemFromMenu
              If FormView = FormViews.Modern Then
                Dim vMenuCommand As New MenuToolbarCommand(vTag, vRow("ExplorerLinkDesc").ToString, 0)
                vMenuCommand.ExplorerMenuAttribute = New ExplorerMenuAttribute(ExplorerMenuSection.Trader, ExplorerMenuCategory.ExplorerMenuLinks)
                vMenuCommand.OnClick = AddressOf ProcessExplorerItemFromCommand
                vMenuCommand.SerialisationFormat = "EXP{0}"
                mvMenuItems.Add(vTag, vMenuCommand)
              End If
            End If
          Next
        End If
      End If
    End If

    LoadWorkstreamMenu(SystemToolStripMenuItem)

    If mvApplicationsMenu.DropDownItems.Count = 0 Then mvApplicationsMenu.Visible = False

    Dim vCheckCount As Integer = 1
    If QueryToolStripMenuItem.DropDownItems.ContainsKey("msiQueryByExample") Then vCheckCount += 1
    If QueryToolStripMenuItem.DropDownItems.Count < vCheckCount Then QueryToolStripMenuItem.Visible = False 'The one item is because of the dummy item used only for the icon
  End Sub

  Friend Sub Execute(pCommand As CommandIndexes)
    Dim pCommandInt As Integer = DirectCast(pCommand, Integer)
    Dim vCommand As MenuToolbarCommand = Me.MenuItems.Cast(Of MenuToolbarCommand).ToList().FirstOrDefault(Function(vItem) vItem.CommandID = pCommandInt)
    If vCommand IsNot Nothing AndAlso vCommand.IsEnabled Then
      vCommand.Click()
    End If
  End Sub

  Private Enum GroupMenuItemTypes
    gmiNewContact
    gmiNewOrganisation
    gmiNewEvent
    gmiFindContact
    gmiFindOrganisation
    gmiFindEvent
    gmiQueryContact
    gmiQueryOrganisation
    gmiQueryEvent
  End Enum

  Private Sub AddGroupMenuItems(ByVal pType As EntityGroup.EntityGroupTypes, ByVal pMenu As ToolStripMenuItem, ByVal pItemType As GroupMenuItemTypes)
    Dim vNewContactIndexes() As CommandIndexes = {CommandIndexes.cbiNewContact, CommandIndexes.cbiNewContact2, CommandIndexes.cbiNewContact3, CommandIndexes.cbiNewContact4, CommandIndexes.cbiNewContact5}
    Dim vNewOrganisationIndexes() As CommandIndexes = {CommandIndexes.cbiNewOrganisation, CommandIndexes.cbiNewOrganisation2, CommandIndexes.cbiNewOrganisation3, CommandIndexes.cbiNewOrganisation4, CommandIndexes.cbiNewOrganisation5}
    Dim vNewEventIndexes() As CommandIndexes = {CommandIndexes.cbiNewEvent, CommandIndexes.cbiNewEvent2, CommandIndexes.cbiNewEvent3, CommandIndexes.cbiNewEvent4, CommandIndexes.cbiNewEvent5}
    Dim vFindContactIndexes() As CommandIndexes = {CommandIndexes.cbiContactFinder, CommandIndexes.cbiContactFinder2, CommandIndexes.cbiContactFinder3, CommandIndexes.cbiContactFinder4, CommandIndexes.cbiContactFinder5}
    Dim vFindOrganisationIndexes() As CommandIndexes = {CommandIndexes.cbiOrganisationFinder, CommandIndexes.cbiOrganisationFinder2, CommandIndexes.cbiOrganisationFinder3, CommandIndexes.cbiOrganisationFinder4, CommandIndexes.cbiOrganisationFinder5}
    Dim vFindEventIndexes() As CommandIndexes = {CommandIndexes.cbiEventFinder, CommandIndexes.cbiEventFinder2, CommandIndexes.cbiEventFinder3, CommandIndexes.cbiEventFinder4, CommandIndexes.cbiEventFinder5}
    Dim vQueryContactIndexes() As CommandIndexes = {CommandIndexes.cbiQueryByExampleContacts, CommandIndexes.cbiQueryByExampleContacts2, CommandIndexes.cbiQueryByExampleContacts3, CommandIndexes.cbiQueryByExampleContacts4, CommandIndexes.cbiQueryByExampleContacts5}
    Dim vQueryOrganisationIndexes() As CommandIndexes = {CommandIndexes.cbiQueryByExampleOrganisations, CommandIndexes.cbiQueryByExampleOrganisations2, CommandIndexes.cbiQueryByExampleOrganisations3, CommandIndexes.cbiQueryByExampleOrganisations4, CommandIndexes.cbiQueryByExampleOrganisations5}
    Dim vQueryEventIndexes() As CommandIndexes = {CommandIndexes.cbiQueryByExampleEvents, CommandIndexes.cbiQueryByExampleEvents2, CommandIndexes.cbiQueryByExampleEvents3, CommandIndexes.cbiQueryByExampleEvents4, CommandIndexes.cbiQueryByExampleEvents5}

    Dim vFinderItems As Integer = 0
    Dim vCommand As MenuToolbarCommand
    Dim vList As CollectionList(Of EntityGroup)
    If pType = EntityGroup.EntityGroupTypes.egtEventGroup Then
      vList = DataHelper.EventGroups
    Else
      vList = DataHelper.ContactAndOrganisationGroups
    End If
    Dim vNewMenuItem As Boolean
    Dim vQueryMenuItem As Boolean
    Dim vCommandIndex As CommandIndexes
    For Each vEntityGroup As EntityGroup In vList
      If vEntityGroup.Type = pType Then
        Select Case pItemType
          Case GroupMenuItemTypes.gmiNewContact
            vCommandIndex = vNewContactIndexes(vFinderItems)
            vNewMenuItem = True
          Case GroupMenuItemTypes.gmiNewOrganisation
            vCommandIndex = vNewOrganisationIndexes(vFinderItems)
            vNewMenuItem = True
          Case GroupMenuItemTypes.gmiNewEvent
            vCommandIndex = vNewEventIndexes(vFinderItems)
            vNewMenuItem = True
          Case GroupMenuItemTypes.gmiFindContact
            vCommandIndex = vFindContactIndexes(vFinderItems)
          Case GroupMenuItemTypes.gmiFindOrganisation
            vCommandIndex = vFindOrganisationIndexes(vFinderItems)
          Case GroupMenuItemTypes.gmiFindEvent
            vCommandIndex = vFindEventIndexes(vFinderItems)
          Case GroupMenuItemTypes.gmiQueryContact
            vCommandIndex = vQueryContactIndexes(vFinderItems)
            vQueryMenuItem = True
          Case GroupMenuItemTypes.gmiQueryOrganisation
            vCommandIndex = vQueryOrganisationIndexes(vFinderItems)
            vQueryMenuItem = True
          Case GroupMenuItemTypes.gmiQueryEvent
            vCommandIndex = vQueryEventIndexes(vFinderItems)
            vQueryMenuItem = True
        End Select
        vCommand = mvMenuItems(vCommandIndex.ToString)
        vCommand.EntityGroup = vEntityGroup
        If vNewMenuItem Then
          vEntityGroup.CanCreate = Not vCommand.HideItem
          vCommand.MenuText = vEntityGroup.GroupName & "..."
          vCommand.ToolTipText = String.Format(ControlText.MnuMNewItem, vEntityGroup.GroupName)
          vCommand.ToolBarText = String.Format(ControlText.MnuMNewItem, vEntityGroup.GroupName) 'BR17806
        ElseIf vQueryMenuItem Then
          vCommand.MenuText = vEntityGroup.GroupDescription & "..."
          vCommand.ToolTipText = String.Format(ControlText.MnuMQueryItem, vEntityGroup.GroupDescription)
          vCommand.ToolBarText = String.Format(ControlText.MnuMQueryItem, vEntityGroup.GroupDescription)  'BR17806
        Else
          vCommand.MenuText = vEntityGroup.GroupDescription & "..."
          vCommand.ToolTipText = String.Format(ControlText.MnuMFindItem, vEntityGroup.GroupDescription)
          vCommand.ToolBarText = String.Format(ControlText.MnuMFindItem, vEntityGroup.GroupDescription) 'BR17806
        End If
        vCommand.AddToMenu(pMenu)
        vFinderItems += 1
        If vFinderItems > 4 Then Exit For
      End If
    Next
    For vFinderItems = vFinderItems To 4
      Select Case pItemType
        Case GroupMenuItemTypes.gmiNewContact
          vCommandIndex = vNewContactIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiNewOrganisation
          vCommandIndex = vNewOrganisationIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiNewEvent
          vCommandIndex = vNewEventIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiFindContact
          vCommandIndex = vFindContactIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiFindOrganisation
          vCommandIndex = vFindOrganisationIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiFindEvent
          vCommandIndex = vFindEventIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiQueryContact
          vCommandIndex = vQueryContactIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiQueryOrganisation
          vCommandIndex = vQueryOrganisationIndexes(vFinderItems)
        Case GroupMenuItemTypes.gmiQueryEvent
          vCommandIndex = vQueryEventIndexes(vFinderItems)
      End Select
      mvMenuItems(vCommandIndex.ToString).HideItem = True
    Next
  End Sub

  Private Sub AddSubMenuItems(ByVal pItems() As CommandIndexes, ByVal pSubMenuText As String, ByVal pParentMenu As ToolStripMenuItem)
    Dim vSubMenu As New ToolStripMenuItem(pSubMenuText)
    For vItem As Integer = 0 To pItems.Length - 1
      mvMenuItems(pItems(vItem).ToString).AddToMenu(vSubMenu)
    Next
    If vSubMenu.DropDownItems.Count > 0 Then pParentMenu.DropDownItems.Add(vSubMenu)
  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCommand As MenuToolbarCommand = Nothing
    If TypeOf sender Is ToolStripItem Then
      Dim vToolStripItem As ToolStripItem = DirectCast(sender, ToolStripItem)
      vCommand = DirectCast(vToolStripItem.Tag, MenuToolbarCommand)
    ElseIf TypeOf sender Is MenuToolbarCommand Then
      vCommand = DirectCast(sender, MenuToolbarCommand)
    End If
    If vCommand IsNot Nothing Then
      ProcessMenuItem(CType(vCommand.CommandID, CommandIndexes))
    End If
  End Sub

  Private Sub WindowMenuOpeningHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vMenu As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
    vMenu.DropDownItems.Clear()
    For Each vForm As Form In MainHelper.Forms
      Dim vItem As New ToolStripMenuItem(vForm.Text, vForm.Icon.ToBitmap, AddressOf WindowMenuHandler)
      vItem.Tag = vForm
      vMenu.DropDownItems.Add(vItem)
    Next
    vMenu.DropDownItems.Add(New ToolStripSeparator)
    vMenu.DropDownItems.Add(mvMenuItems(CommandIndexes.cbiCloseAll.ToString).MenuStripItem)
  End Sub

  Private Sub QueryMenuOpeningHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vMenu As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
    If vMenu.DropDownItems.ContainsKey("msiQueryByExample") Then
      vMenu.DropDownItems("msiQueryByExample").Visible = False      'Hide this as the menu is only there to support the icon
    End If
  End Sub

  Private Sub SystemMenuOpeningHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vMenu As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
    For Each vMenuItem As ToolStripMenuItem In vMenu.DropDownItems
      If vMenuItem.DropDownItems.ContainsKey("msiPostcodeValidation") Then
        vMenuItem.DropDownItems("msiPostcodeValidation").Visible = AppValues.IsPAFSupported
      End If
    Next
  End Sub

  Private Sub WindowMenuHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vForm As Form = DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, Form)
    vForm.BringToFront()
  End Sub

  Private Sub ProcessMenuItem(ByVal pCommand As CommandIndexes)
    Dim vCursor As New BusyCursor
    Try
      Select Case pCommand
        Case CommandIndexes.cbiLogin
          Application.Restart()
        Case CommandIndexes.cbiNewAction
          FormHelper.EditAction(0)
        Case CommandIndexes.cbiNewActionTemplate
          FormHelper.EditActionTemplate(0)
        Case CommandIndexes.cbiNewContact, CommandIndexes.cbiNewContact2, CommandIndexes.cbiNewContact3, CommandIndexes.cbiNewContact4, CommandIndexes.cbiNewContact5
          Dim vList As New ParameterList
          vList("ContactGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctContact, vList)
        Case CommandIndexes.cbiNewDocument
          If MainHelper.CurrentContact IsNot Nothing AndAlso MainHelper.CurrentContact.ContactNumber > 0 Then
            'BR21352 Update the Current ContactInfo, we don't know what has happened in CardMaintenance
            MainHelper.SetStatusContact(New ContactInfo(MainHelper.CurrentContact.ContactNumber), True)
          End If
          FormHelper.NewDocument()
        Case CommandIndexes.cbiNewOrganisation, CommandIndexes.cbiNewOrganisation2, CommandIndexes.cbiNewOrganisation3, CommandIndexes.cbiNewOrganisation4, CommandIndexes.cbiNewOrganisation5
          Dim vList As New ParameterList
          vList("OrganisationGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctOrganisation, vList)
        Case CommandIndexes.cbiNewEvent, CommandIndexes.cbiNewEvent2, CommandIndexes.cbiNewEvent3, CommandIndexes.cbiNewEvent4, CommandIndexes.cbiNewEvent5
          Dim vList As New ParameterList
          vList("EventGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowEventIndex(0, mvMenuItems(pCommand.ToString).EntityGroup.Code)
        Case CommandIndexes.cbiNewTelephoneCall
          Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctTCRDocument)
          AddHandler vForm.SelectedContactChanged, AddressOf vForm_SelectedContactChanged
          vForm.Show()
        Case CommandIndexes.cbiNewSelectionSet
          FormHelper.AddSelectionSet(mvParentForm)
        Case CommandIndexes.cbiPreferences
          Dim vForm As Form
          If Not MainHelper.MainForm.IsMdiContainer Then
            vForm = New frmPreferences
            vForm.ShowDialog(CurrentMainForm)
          Else
            For Each vForm In MainHelper.Forms
              If TypeOf (vForm) Is frmPreferences Then
                vForm.BringToFront()
                Return
              End If
            Next
            vForm = New frmPreferences
            vForm.Show()
          End If

        Case CommandIndexes.cbiLogWEBServices
          AppValues.LogWEBServiceCalls = Not AppValues.LogWEBServiceCalls
          mvMenuItems(pCommand.ToString).CheckToolStripItem(mvMenuStrip, Nothing, AppValues.LogWEBServiceCalls)
        Case CommandIndexes.cbiPageSetup
          Dim vDoc As New PrintDocument
          vDoc.DefaultPageSettings = PrintHandler.DefaultPageSettings
          vDoc.PrinterSettings = PrintHandler.DefaultPrinterSettings
          Dim vPsd As New PageSetupDialog
          vPsd.Document = vDoc
          vPsd.AllowMargins = True
          vPsd.EnableMetric = True
          vPsd.ShowDialog(mvParentForm)
        Case CommandIndexes.cbiExit
          MainHelper.CloseAllForms()
        Case CommandIndexes.cbiNextRecord
          MainHelper.NavigateToNext()
        Case CommandIndexes.cbiPreviousRecord
          MainHelper.NavigateToPrevious()
        Case CommandIndexes.cbiToolbar
          MainHelper.ShowToolbar = Not MainHelper.ShowToolbar
        Case CommandIndexes.cbiNavigationPanel
          MainHelper.NavigationPanel = Not MainHelper.NavigationPanel
        Case CommandIndexes.cbiStatusBar
          MainHelper.StatusBar = Not MainHelper.StatusBar
        Case CommandIndexes.cbiHeaderPanel
          MainHelper.ShowHeaderPanel = Not MainHelper.ShowHeaderPanel
        Case CommandIndexes.cbiSelectionPanel
          MainHelper.ShowSelectionPanel = Not MainHelper.ShowSelectionPanel
        Case CommandIndexes.cbiDashboard
          Dim vForm As Form
          For Each vForm In MainHelper.Forms
            If TypeOf (vForm) Is frmDashboard Then
              If vForm.Visible = False Then
                vForm.Close()
              Else
                vForm.BringToFront()
                Return
              End If
            End If
          Next
          vForm = New frmDashboard
          FormHelper.InitWindowForViewType(vForm)
          vForm.Show()
        Case CommandIndexes.cbiMyDetails
          FormHelper.ShowContactCardIndex(DataHelper.UserInfo.ContactNumber)
        Case CommandIndexes.cbiMyOrganisation
          FormHelper.ShowContactCardIndex(DataHelper.UserInfo.OrganisationNumber)
        Case CommandIndexes.cbiMyActions
          Dim vList As New ParameterList
          vList("MyActions") = "Y"
          vList("ActionStatus") = AppValues.ActiveActionStatus
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftActions, vList)
        Case CommandIndexes.cbiMyDocuments
          Dim vList As New ParameterList
          vList("Outstanding") = "Y"
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDocuments, vList)
        Case CommandIndexes.cbiMyInBox
          If EMailApplication.EmailInterface.CanEMail(True) Then
            Dim vInBox As New frmInbox
            vInBox.Show()
          Else
            ShowInformationMessage(InformationMessages.ImEMailNotConfigured, AppValues.ConfigurationValue(AppValues.ConfigurationValues.email_interface))
          End If
        Case CommandIndexes.cbiMyJournal
          FormHelper.ShowCardDisplay(CareServices.XMLContactDataSelectionTypes.xcdtContactJournals, 0)
        Case CommandIndexes.cbiRefresh
          For Each vForm As Form In MainHelper.Forms
            If TypeOf (vForm) Is MaintenanceParent Then
              DirectCast(vForm, MaintenanceParent).RefreshData()
            ElseIf TypeOf (vForm) Is frmCampaignSet Then
              DirectCast(vForm, frmCampaignSet).RefreshData()
            End If
          Next
        Case CommandIndexes.cbiListManager
          Dim vLM As New ListManager(0, False)
          vLM.Show()
        Case CommandIndexes.cbiDocumentDistributor
          Dim vForm As New frmDocumentDistributor
          If vForm.InitDocumentDistributor() Then
            vForm.Show()
          Else
            vForm.Dispose()
          End If
        Case CommandIndexes.cbiSendEmail
          If EMailApplication.EmailInterface.CanEMail Then
            If MainHelper.CurrentContact IsNot Nothing AndAlso MainHelper.CurrentContact.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
              Dim vForm As frmSelectItems
              vForm = New frmSelectItems(MainHelper.CurrentContact)
              If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                EMailApplication.EmailInterface.SendMail(mvParentForm, EmailInterface.SendEmailOptions.seoAddressResolveUI Or EmailInterface.SendEmailOptions.seoMultipleRecipients, "", "", vForm.EMailAddresses)
              End If
            Else
              Dim vEMailAddress As String = ""
              If MainHelper.CurrentContact IsNot Nothing Then vEMailAddress = MainHelper.CurrentContact.EMailAddresses
              EMailApplication.EmailInterface.SendMail(mvParentForm, EmailInterface.SendEmailOptions.seoAddressResolveUI Or EmailInterface.SendEmailOptions.seoMultipleRecipients, "", "", vEMailAddress)
            End If
          End If
        Case CommandIndexes.cbiCustomise
          Dim vForm As Form
          vForm = New frmCustomiseToolBar(AppHelper.ImageProvider.NewImageList32, AppHelper.ImageProvider.NewImageList32, mvMenuItems, mvToolStrip, CInt(CommandIndexes.cbiSeparator))
          vForm.ShowDialog()

        Case CommandIndexes.cbiCloseOpenBatch
          FormHelper.ShowBatchFinder(0, , mvParentForm, True, Nothing)
        Case CommandIndexes.cbiCopyEventPricingMatrix
          CopyEventPricingMatrix()
        Case CommandIndexes.cbiClearCache
          DataHelper.ClearCachedData()
          ExamsDataHelper.ClearCachedData()
          ObjectCache.ClearAllCaches()
        Case CommandIndexes.cbiAllowCaching
          AppValues.AllowCaching = Not AppValues.AllowCaching
          mvMenuItems(CommandIndexes.cbiAllowCaching.ToString()).CheckToolStripItem(mvMenuStrip, Nothing, AppValues.AllowCaching)
          ProcessMenuItem(CommandIndexes.cbiClearCache)
        Case CommandIndexes.cbiFundraisingPaymentScheduleFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftFundraisingPaymentScheduleFinder)
        Case CommandIndexes.cbiFundraisingRequestFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftFundraisingRequestsFinder)
        Case CommandIndexes.cbiRunMailing
          FormHelper.RunMailing(CareNetServices.TaskJobTypes.tjtSelectMailing)
        Case CommandIndexes.cbiRunReport
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "USER"
          Dim vDataSet As DataSet = DataHelper.GetLookupDataSet(CareNetServices.XMLLookupDataTypes.xldtReports, vList)
          Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litUserReports)
          If vForm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim vReportNumber As Integer = IntegerValue(DataHelper.GetTableFromDataSet(vDataSet).Rows(vForm.SelectedRow)("ReportNumber").ToString)
            If vReportNumber > 0 Then
              Dim vParameterForm As New frmApplicationParameters(vReportNumber, "USER", Nothing, CareNetServices.FunctionParameterTypes.fptNone)
              Dim vReportList As ParameterList = Nothing
              If vParameterForm.HasControls Then
                If vParameterForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                  vReportList = vParameterForm.ReportParameterList
                End If
              Else
                vReportList = New ParameterList(True)
              End If
              If vReportList IsNot Nothing Then
                vReportList("ReportCode") = "USER"
                vReportList.IntegerValue("ReportNumber") = vReportNumber
                Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
              End If
            End If
          End If
        Case CommandIndexes.cbiSearchData
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftTextSearch)
        Case CommandIndexes.cbiQueryByExample
          Dim vList As New ParameterList
          Dim vFinderType As CareNetServices.XMLDataFinderTypes
          If TypeOf (mvParentForm) Is frmCardSet Then
            Dim vCardSet As frmCardSet = DirectCast(mvParentForm, frmCardSet)
            If vCardSet.ContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
              vFinderType = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleOrganisations
              vList("OrganisationGroup") = vCardSet.ContactInfo.ContactTypeCode
            Else
              vFinderType = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts
              vList("ContactGroup") = vCardSet.ContactInfo.ContactTypeCode
            End If
          ElseIf TypeOf (mvParentForm) Is frmEventSet Then
            vFinderType = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents
            Dim vEventSet As frmEventSet = DirectCast(mvParentForm, frmEventSet)
            vList("EventGroup") = vEventSet.CareEventInfo.EventGroup
          Else
            vFinderType = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts
            vList("ContactGroup") = "CON"
          End If
          FormHelper.ShowFinder(CType(vFinderType, CareServices.XMLDataFinderTypes), vList)
        Case CommandIndexes.cbiQueryByExampleContacts, CommandIndexes.cbiQueryByExampleContacts2, CommandIndexes.cbiQueryByExampleContacts3, CommandIndexes.cbiQueryByExampleContacts4, CommandIndexes.cbiQueryByExampleContacts5
          Dim vFinderType As CareNetServices.XMLDataFinderTypes = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts
          Dim vList As New ParameterList
          vList("ContactGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowFinder(CType(vFinderType, CareServices.XMLDataFinderTypes), vList)
        Case CommandIndexes.cbiQueryByExampleOrganisations, CommandIndexes.cbiQueryByExampleOrganisations2, CommandIndexes.cbiQueryByExampleOrganisations3, CommandIndexes.cbiQueryByExampleOrganisations4, CommandIndexes.cbiQueryByExampleOrganisations5
          Dim vFinderType As CareNetServices.XMLDataFinderTypes = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleOrganisations
          Dim vList As New ParameterList
          vList("OrganisationGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowFinder(CType(vFinderType, CareServices.XMLDataFinderTypes), vList)
        Case CommandIndexes.cbiQueryByExampleEvents, CommandIndexes.cbiQueryByExampleEvents2, CommandIndexes.cbiQueryByExampleEvents3, CommandIndexes.cbiQueryByExampleEvents4, CommandIndexes.cbiQueryByExampleEvents5
          Dim vFinderType As CareNetServices.XMLDataFinderTypes = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents
          Dim vList As New ParameterList
          vList("EventGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowFinder(CType(vFinderType, CareServices.XMLDataFinderTypes), vList)

        Case CommandIndexes.cbiContactFinder, CommandIndexes.cbiContactFinder2, CommandIndexes.cbiContactFinder3, CommandIndexes.cbiContactFinder4, CommandIndexes.cbiContactFinder5
          Dim vList As New ParameterList
          If pCommand <> CommandIndexes.cbiContactFinder Then vList("LockContactGroup") = "Y"
          vList("ContactGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftContacts, vList)
        Case CommandIndexes.cbiOrganisationFinder, CommandIndexes.cbiOrganisationFinder2, CommandIndexes.cbiOrganisationFinder3, CommandIndexes.cbiOrganisationFinder4, CommandIndexes.cbiOrganisationFinder5
          Dim vList As New ParameterList
          If pCommand <> CommandIndexes.cbiOrganisationFinder Then vList("LockContactGroup") = "Y"
          vList("OrganisationGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          vList("IncludeGroupsFromContactCard") = "Y"
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftOrganisations, vList)
        Case CommandIndexes.cbiDocumentFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDocuments)
        Case CommandIndexes.cbiMeeting
          FormHelper.ShowFinder(CType(CareNetServices.XMLDataFinderTypes.xdftMeetings, CareServices.XMLDataFinderTypes))
        Case CommandIndexes.cbiActionFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftActions)
        Case CommandIndexes.cbiSelectionSetFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftSelectionSets)
        Case CommandIndexes.cbiEventFinder, CommandIndexes.cbiEventFinder2, CommandIndexes.cbiEventFinder3, CommandIndexes.cbiEventFinder4, CommandIndexes.cbiEventFinder5
          Dim vList As New ParameterList
          vList("EventGroup") = mvMenuItems(pCommand.ToString).EntityGroup.Code
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftEvents, vList)
        Case CommandIndexes.cbiMemberFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftMembers)
        Case CommandIndexes.cbiPayPlanFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftPaymentPlans)
        Case CommandIndexes.cbiCovenantFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftCovenants)
        Case CommandIndexes.cbiXactionFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftTransactions)
        Case CommandIndexes.cbiStandingOrderFinder
          Dim vList As New ParameterList
          vList("AllowReconcile") = "N"
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftStandingOrders, vList)
        Case CommandIndexes.cbiDirectDebitFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDirectDebits)
        Case CommandIndexes.cbiCCCAFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftCreditCardAuthorities)
        Case CommandIndexes.cbiGADFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftGiftAidDeclarations)
        Case CommandIndexes.cbiInvoiceFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftInvoiceCreditNotes)
        Case CommandIndexes.cbiLegacyFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftLegacies)
        Case CommandIndexes.cbiGAYEFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftPayrollGiving)
        Case CommandIndexes.cbiPostTaxPGFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftPostTaxPayrollGiving)
        Case CommandIndexes.cbiPurchaseOrderFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftPurchaseOrders)
        Case CommandIndexes.cbiProductFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftProducts)
        Case CommandIndexes.cbiCampaignFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftCampaigns)
        Case CommandIndexes.cbiStandardDocument
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftStandardDocuments)
        Case CommandIndexes.cbiServiceProductFinder
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftServiceProducts)
        Case CommandIndexes.mnuAdminAmendmentHistory
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtAmendmentHistoryView, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminDeleteDormantContacts
          NYI()
        Case CommandIndexes.mnuAdminProcessAddressChanges
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtProcessAddressChanges, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminSetPostDatedContacts
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtSetPostDatedContacts, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminUpdateActionStatuses
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUpdateActionStatus, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminUpdateRegionalData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGenerateAddressGeoRegions, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminUpdateMailsort
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtMailsortUpdate, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminUpdatePrincipalUser
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUpdatePrincipalUser, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case CommandIndexes.mnuAdminUpdateSearchNames
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUpdateSearchNames, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminPostcodeValidation
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPostcodeValidation, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminPurgeStickyNotes
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtPurgeStickyNotes, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDeDuplicationContactMerge
          FormHelper.DoContactMerge(False)
        Case CommandIndexes.mnuDeDuplicationAddressMerge
          FormHelper.DoAddressMerge()
        Case CommandIndexes.mnuDeDuplicationOrganisationMerge
          FormHelper.DoContactMerge(True)
        Case CommandIndexes.mnuDeDuplicationAmalgamateOrganisations
          FormHelper.DoAmalgamateOrganisation()
        Case CommandIndexes.mnuDeDuplicationContactDeDuplication
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtContactDeDuplication, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDeDuplicationBulkAddressMerge
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBulkAddressMerge, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDeDuplicationBulkContactMerge
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBulkMerge, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDeDuplicationBulkOrganisationMerge
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBulkOrganisationMerge, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDeDuplicateProcessDuplicateContacts
          FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftDuplicateContactRecords)
        Case CommandIndexes.mnuCMProcessMetaData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCheetahMailMetaData, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuCMProcessEventData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCheetahMailEventData, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuCMProcessTotals
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCheetahMailTotals, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuCMUpdateBulkMailer
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBulkMailerStatistics, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuFinBViewBatchDetail
          Dim vBatchNumber As Integer
          Dim vFinderList As New ParameterList
          vFinderList("AllowNew") = "N"
          If FormHelper.ShowBatchFinder(vBatchNumber, vFinderList, mvParentForm) Then
            FormHelper.ShowViewBatchDetails(vBatchNumber)
          End If
        Case CommandIndexes.mnuFinBProcessBatches
          Dim vForm As New frmBatchProcessing
          vForm.Show()
        Case CommandIndexes.mnuFinBOutstandingBatchesReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtOutstandingBatchesReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinBChequeList
          Dim vBatchNumber As Integer
          Dim vList As New ParameterList
          vList("AllowNew") = "N"
          vList("SuppressBatchView") = "Y"
          FormHelper.ShowBatchFinder(vBatchNumber, vList, mvParentForm)
          If vBatchNumber > 0 Then
            Dim vBatchProcess As New frmBatchProcessing
            vBatchProcess.PrintChequeList(vBatchNumber)
          End If
        Case CommandIndexes.mnuFinBPayingInSlips
          Dim vBatchProcess As New frmBatchProcessing
          vBatchProcess.ProcessPIS()
        Case CommandIndexes.mnuFinBCashBookBatch
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCashBookPosting, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinBUpdateBatch
          Dim vBatchNumber As Integer
          Dim vList As New ParameterList
          vList("AllowNew") = "N"
          vList("SuppressBatchView") = "Y"
          FormHelper.ShowBatchFinder(vBatchNumber, vList, mvParentForm)
          If vBatchNumber > 0 Then
            Dim vParams As New ParameterList(True)
            vParams("BatchNumber") = vBatchNumber.ToString
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBatchUpdate, vParams, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
          End If
        Case CommandIndexes.mnuFinBCreateJournal
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCreateJournalFiles, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinBSummaryReport
          RunBatchReport("FPBSUM")
        Case CommandIndexes.mnuFinBDetailReport
          RunBatchReport("FPBDET")
        Case CommandIndexes.mnuFinBPurgeOldBatches
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBatchPurge, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinBPurgePrizeDrawBatches
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPurgePrizeDrawBatches, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDMandateFile
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDDMandateFile, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDBatches
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDirectDebitRun, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDClaimFile
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDDClaimFile, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDDirectCreditFile
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDDCreditFile, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDBACSRejections
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBACSRejections, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDUploadBacsMessagingData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUploadBACSMessagingData, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinDDConvertManualDirectDebits
          If ShowQuestion(QuestionMessages.QmWarningManualDirectDebit, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtConvertManualDirectDebits, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          End If
        Case CommandIndexes.mnuFinCCBatches
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCreditCardRun, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCCCreateFile
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCCClaimFile, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCCCreateReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCCClaimReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCCCreateCardSalesFile
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCardSalesFile, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCCCreateCardSalesReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCardSalesReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCCAuthorisationsReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCreditCardAuthorisationReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuFinSOLoadStatement
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtStatementLoader, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSOReconciliation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtAutoSOReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSOCancellation
          'FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtStandingOrderCancellation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyStandingOrderCancellation, CareServices.TaskJobTypes.tjtStandingOrderCancellation)
          vGenMail.Process(0)
        Case CommandIndexes.mnuFinSOManualRec
          '<TODO:Deferred>
          Dim vParams As New ParameterList(True)
          vParams("ManualSOFinder") = "Y"
          If (FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftStandingOrders, vParams) > 0) Then
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtManualSOReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          End If
        Case CommandIndexes.mnuFinSOReconReport
          Dim vReportCode As String = ""
          'vReportCode = "PSOR"
          'BR18053-Amended the parameters passed to frmApplicationParameters, as passing PSOR as ReportCode was incorrect.
          Dim vAP As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptSOReconciliationReport, New ParameterList, Nothing)
          If vAP.HasControls Then
            Dim vDialogResult As DialogResult = vAP.ShowDialog()
            Dim vReportList As ParameterList = Nothing
            If vDialogResult = System.Windows.Forms.DialogResult.OK Then
              vReportList = vAP.ReturnList
              If vReportList.Contains("Date") Then vReportList("RP1") = vReportList("Date")
              If vReportList.Contains("ReconciledStatus") Then vReportList("RP2") = vReportList("ReconciledStatus")
              If vReportList.Contains("BankAccount") Then vReportList("RP3") = vReportList("BankAccount")

              Select Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_statement_input_format, "")
                Case "BANKLINE"
                  vReportList("ReportCode") = "SOREBL"
                  Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
                Case Else
                  vReportList("ReportCode") = "SORECO"
                  Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
              End Select
            End If
          End If
        Case CommandIndexes.mnuFinSOBankTransactionsReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBankTransactionsReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPRenewals
          '<TODO: - Count>          
          Dim vDefaults As New ParameterList()
          vDefaults("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtRenewalsAndReminders, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPRemoveArrears
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtRemoveArrears, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPExpiry
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtExpirePaymentPlans, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPNonMemberFulfilment
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyNonMemberFulfilment, CareServices.TaskJobTypes.tjtMemberFulfilment)
          vGenMail.Process(0)
        Case CommandIndexes.mnuFinPPUpdateProducts
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUpdatePaymentPlanProducts, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPApplySurcharges
          Dim vDefaults As New ParameterList(True)
          vDefaults("IgnorePartPayments") = CBoolYN(BooleanValue(AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.def_surcharge_ignore_part_pay)))
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtApplyPaymentPlanSurcharges, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPReCalcLoanInterest
          Dim vDefaults As New ParameterList(True)
          vDefaults("CalculationDate") = Today.ToShortDateString()
          vDefaults("MinDate") = Today.AddYears(-1).ToShortDateString
          vDefaults("MaxDate") = Today.AddYears(1).ToShortDateString
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtRecalculateLoanInterest, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPUpdateLoanInterestRates
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtUpdateLoanInterestRates, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPPTransferPaymentPlanChanges
          Dim vDefaults As New ParameterList(True)
          vDefaults("FromDate") = AppValues.TodaysDate
          vDefaults("ToDate") = AppValues.TodaysDate
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtTransferPaymentPlanChanges, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTPickingLists
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPickingList, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTConfirmAllocation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtConfirmStockAllocation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTAllocateToBO
          Dim vList As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBackOrderAllocation, vList, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTDespatchNotes
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDespatchNotes, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTDespatchTracking
          Dim vForm As frmDespatchTracking = New frmDespatchTracking()
          vForm.ShowDialog()
        Case CommandIndexes.mnuFinSTMovement
          Dim vForm As frmStockMovement = New frmStockMovement()
          vForm.ShowDialog()
        Case CommandIndexes.mnuFinSTTransferStockToPack
          Dim vForm As frmTransferStockToPack = New frmTransferStockToPack()
          vForm.ShowDialog()
        Case CommandIndexes.mnuFinSTExport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtStockExport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTPurgeBackOrders
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBackOrderPurge, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTPurgePickingAndDespatch
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPickingAndDespatchPurge, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinSTPLAwaitingConfirm
          Dim vDataSet As DataSet = DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstSelectAwaitListConfirmation)
          Dim vForm As frmSelectItems = New frmSelectItems(vDataSet, frmSelectItems.SelectItemsTypes.sitSelectAwaitListConfirmation)
          vForm.ShowDialog()
        Case CommandIndexes.mnuFinSTValuationReport
          Dim vReportList As New ParameterList(True, True)
          vReportList("ReportCode") = "STKVAL"
          Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case CommandIndexes.mnuFinCSTransferInvoices
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtInvoiceTransfer, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCSTransferCustomers
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCustomerTransfer, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCSStatementGeneration
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtCreditStatementGeneration, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case CommandIndexes.mnuFinPRPriceChange
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPriceChange, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPRPurchasedProductReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPurchasedProductReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGADConfirmation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGADConfirmation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGADPotentialClaim
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGiftAidPotentialClaim, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGADTaxClaim
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGiftAidClaim, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGADGiftDetails
          Dim vReportCode As String = String.Empty
          Select Case AppValues.ControlValue(AppValues.ControlTables.gift_aid_controls, AppValues.ControlValues.claim_file_format).ToUpper
            Case "C"
              vReportCode = "GACCSV"
            Case "O"
              vReportCode = "GAPORT"
            Case Else   'F
              vReportCode = "GACSUM"
          End Select

          Dim vReportList As New ParameterList(True, True)
          If AppValues.ControlValue(AppValues.ControlTables.gift_aid_controls, AppValues.ControlValues.claim_file_format).ToUpper.Equals("O") Then
            vReportList("RP2") = AppValues.ControlValue(AppValues.ControlTables.gift_aid_controls, AppValues.ControlValues.submitter_contact)
            vReportList("RPpAdjustmentText") = AppValues.ControlValue(AppValues.ControlTables.gift_aid_controls, AppValues.ControlValues.adjustment_text)
          Else
            vReportList("RPReportType") = "ActualClaim"
            vReportList("RPPaymentsFrom") = Convert.ToString(Date.Today)
            vReportList("RPPaymentsTo") = Convert.ToString(Date.Today)
            vReportList("RP2") = "declaration_tax_claim_lines"
            vReportList("RP3") = "Y"
          End If
          vReportList("ReportCode") = vReportCode
          RunClaimReport(vReportList)
        Case CommandIndexes.mnuFinGADGiftAnalysis
          Dim vReportList As New ParameterList(True, True)
          vReportList("ReportCode") = "GACANA"
          RunClaimReport(vReportList)
        Case CommandIndexes.mnuFinBulkGiftAidUpdate
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBulkGiftAidUpdate, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGASPotentialClaim
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGASPotentialClaim, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGASTaxClaim
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGASTaxClaim, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGASClaimDetails
          Dim vReportList As New ParameterList(True, True)
          vReportList("RP2") = "ga_sponsorship_tax_claim_lines"
          vReportList("RP3") = "Y"
          vReportList("RPReportType") = "ActualClaim"
          vReportList("RPPaymentsFrom") = Convert.ToString(Date.Today)
          vReportList("RPPaymentsTo") = Convert.ToString(Date.Today)
          vReportList("ReportCode") = "GASP"
          RunClaimReport(vReportList)
        Case CommandIndexes.mnuFinancialIncentives
          'NYI()
          Dim vForm As frmIncentiveMaintenance = New frmIncentiveMaintenance()
          vForm.ShowDialog()

        Case CommandIndexes.mnuFinGAYELoadPayments
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGAYEPaymentLoader, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGAYEReconciliation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGAYEReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGAYEBulkCancellation
          'NYI()
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyGAYECancellation, CareServices.TaskJobTypes.tjtPayrollPledgeCancellation)
          vGenMail.Process(0)
        Case CommandIndexes.mnuFinGAYEPostTaxPGReconciliation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPostTaxPGReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinBanksLoad
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBankDataLoad, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuFinCAFExpectedPayments
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFExpectedPaymentsReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCAFProvisionalBatchClaim
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFProvisionalBatchClaim, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCAFCreateCardSalesReport
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFCardSalesReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCAFLoadPaymentData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFPaymentLoader, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinCAFReconcilePaymentData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCAFPaymentReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuFinNomSummaryReport, CommandIndexes.mnuFinNomDetailReport
          Dim vReportList As New ParameterList(True, True)
          Select Case pCommand
            Case CommandIndexes.mnuFinNomSummaryReport
              vReportList("RP1") = "Y"
            Case Else
              vReportList("RP2") = "Y"
          End Select
          vReportList("ReportCode") = "NOMCDS"
          Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case CommandIndexes.mnuFinPOTransPayments
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPostPayments, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPOTransSuppliers
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPOTransferSuppliers, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPOAuthorisePayments
          Dim vForm As New frmSelectItems(Nothing, frmSelectItems.SelectItemsTypes.sitAuthorisePOPayments)
          vForm.ShowDialog(mvParentForm)
        Case CommandIndexes.mnuFinPOProcessPayments
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtProcessPurchaseOrderPayments, New ParameterList(True), True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPOAutoGenerate
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CType(CareNetServices.TaskJobTypes.tjtPurchaseOrderGeneration, CareServices.TaskJobTypes), vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinChequeProduction
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtChequeProduction, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPOPrint
          FormHelper.ProcessTask(CType(CareNetServices.TaskJobTypes.tjtPurchaseOrderPrint, CareServices.TaskJobTypes), Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuFinPISLoadStatement
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPISStatementLoader, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinPISReconciliation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPISReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuMemAssumedVotingRights
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtAssumedVotingRights, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemBallotPaperProduction
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBallotPaperProduction, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemBranchDonations
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBranchDonationsReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemBranchIncome
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtBranchIncomeReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemFulfilment
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyMemberFulfilment, CareServices.TaskJobTypes.tjtMemberFulfilment)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMemFutureChanges
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtFutureMembershipChanges, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemGenerateStatistics
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPeriodStatsGenerateData, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemJuniorAnalysis
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtJuniorMembershipAnalysisReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemStatisticsDetailed, CommandIndexes.mnuMemStatisticsSummary
          If pCommand = CommandIndexes.mnuMemStatisticsSummary Then
            Dim vParams As New ParameterList(True)
            vParams("SummaryReport") = "Y"
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPeriodStatsReport, vParams, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          Else
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPeriodStatsReport, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          End If
        Case CommandIndexes.mnuMemSuspension
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtMembershipSuspension, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMemCards
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyMembershipCards, CareServices.TaskJobTypes.tjtMembCardMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuFinGAIClaimDetails
          Dim vReportList As New ParameterList(True, True)
          vReportList("ReportCode") = "IGAT"
          RunClaimReport(vReportList)
        Case CommandIndexes.mnuFinGAIPotentialClaim
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtIrishGiftAidPotentialClaim, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuFinGAITaxClaim
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtIrishGiftAidTaxClaim, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDutchLoadPayments
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDutchElectronicPaymentsLoader, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuDutchProcessPayments
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDutchElectronicPaymentsReconciliation, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuMktGenerateData
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGenerateMarketingData, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuMailingDocsFind
          Dim vFinder As New frmFinder(CareServices.XMLDataFinderTypes.xdftContactMailingDocuments, Nothing)
          vFinder.ShowDialog()
        Case CommandIndexes.mnuMailingDocsProduce
          Dim vMD As New MailingDocument
          vMD.RunMailingDocumentProduction()
        Case CommandIndexes.mnuMailingTYL
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtThankYouLetters, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMailingFinder
          Dim vFinder As New frmFinder(CareServices.XMLDataFinderTypes.xdftMailings, Nothing)
          vFinder.ShowDialog()
        Case CommandIndexes.mnuEMailProcessor
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtEMailProcessor, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMailingListAllContacts
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtListAllContacts, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuMailingIrishGiftAid
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyIrishGiftAid, CareServices.TaskJobTypes.tjtIrishGiftAidMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingMembers
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyMembers, CareServices.TaskJobTypes.tjtMemberMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingPayers
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyPayers, CareServices.TaskJobTypes.tjtPayerMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingPreTextPayrollGivingPledges
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyGAYEPledges, CareServices.TaskJobTypes.tjtPayrollPledgeMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingStandingOrders
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyStandingOrders, CareServices.TaskJobTypes.tjtStandingOrderMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingEventsBookings
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyEventBookings, CareServices.TaskJobTypes.tjtEventBookerMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingEventsDelegates
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyEventAttendees, CareServices.TaskJobTypes.tjtEventDelegateMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingEventsPersonnel
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyEventPersonnel, CareServices.TaskJobTypes.tjtEventPersonnelMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingEventsSponsors
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyEventSponsors, CareServices.TaskJobTypes.tjtEventSponsorMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingExamsBookings
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyExamBookings, CareServices.TaskJobTypes.tjtExamBookerMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingExamsCandidates
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyExamCandidates, CareServices.TaskJobTypes.tjtExamCandidateMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingDirectDebit
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyDirectDebits, CareServices.TaskJobTypes.tjtDirectDebitMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingDirectDebit
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyDirectDebits, CareServices.TaskJobTypes.tjtDirectDebitMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingSubscriptions
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtySubscriptions, CareServices.TaskJobTypes.tjtSubscriptionMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingSelectionTester
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtySelectionTester, CareServices.TaskJobTypes.tjtSelectionTester)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingSelectionManager
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyGeneralMailing, CareServices.TaskJobTypes.tjtSelectionManagerMailing)
          vGenMail.Process(0)
        Case CommandIndexes.mnuMailingDirectDebit
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyDirectDebits, CareServices.TaskJobTypes.tjtDirectDebitMailing)
          vGenMail.Process(0)
          'Case CommandIndexes.mnuMailingDirectDebit
          '  Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyDirectDebits, CareServices.TaskJobTypes.tjtDirectDebitMailing)
          '  vGenMail.Process(0)
        Case CommandIndexes.mnuFinDBRepOpenBoxes, CommandIndexes.mnuFinDBRepUnAllocatedDonations, CommandIndexes.mnuFinDBRepClosedByLocation
          'These reports do not have any parameters
          Dim vReportCode As String = ""
          Select Case pCommand
            Case CommandIndexes.mnuFinDBRepOpenBoxes
              vReportCode = "DBOL"
            Case CommandIndexes.mnuFinDBRepUnAllocatedDonations
              vReportCode = "DBNA"
            Case CommandIndexes.mnuFinDBRepClosedByLocation
              vReportCode = "DBCL"
          End Select
          If vReportCode.Length > 0 Then
            Dim vReportList As New ParameterList(True, True)
            vReportList("ReportCode") = vReportCode
            Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
          End If
        Case CommandIndexes.mnuFinDBRepRollOfHonour
          'This report does not have any parameters but as it is a CSV report it doesn't make sense to Preview it so remove the Preview option
          'Also as it contains multi-line fields they don't display correctly in the preview window
          Dim vOutput As New frmReportOutput(False, False)
          If vOutput.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim vList As New ParameterList(True, True)
            vList("ReportDestination") = vOutput.ReportDestination
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGenerateRollOfHonour, vList, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          End If
        Case CommandIndexes.mnuFinSTBackOrdersReport
          Dim vReportList As New ParameterList(True, True)
          vReportList("ReportCode") = "BAKORD"
          Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)

        Case CommandIndexes.mnuFinSTSalesAnalysis
          Dim vAP As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptStockSalesAnalysis, Nothing, Nothing)
          If vAP.HasControls Then
            Dim vDialogResult As DialogResult = vAP.ShowDialog()
            Dim vReportList As ParameterList = Nothing
            If vDialogResult = System.Windows.Forms.DialogResult.OK Then
              vReportList = vAP.ReturnList
              vReportList("RP2") = Convert.ToString(DateSerial(Year(Date.Today), Month(Date.Today), 1))
              If vReportList.Contains("Date") Then vReportList("RP1") = vReportList("Date")
              If vReportList.Contains("NominalAccount") Then vReportList("RP3") = vReportList("NominalAccount")
              If vReportList.Contains("SalesGroup") Then vReportList("RP4") = vReportList("SalesGroup")
              If vReportList.Contains("Company") Then vReportList("RP5") = vReportList("Company")
              vReportList("ReportCode") = "SSALEA"
              Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
            End If
          End If

        Case CommandIndexes.mnuFinSTSalesAnalysisDetailed, CommandIndexes.mnuFinSTSalesAnalysisSummary
          Dim vReportCode As String = "SSALED"
          Dim vFPType As CareNetServices.FunctionParameterTypes = CareNetServices.FunctionParameterTypes.fptStockSalesAnalysisDetailed
          If pCommand = CommandIndexes.mnuFinSTSalesAnalysisSummary Then
            vReportCode = "SSALES"
            vFPType = CareNetServices.FunctionParameterTypes.fptStockSalesAnalysisSummary
          End If

          Dim vAP As New frmApplicationParameters(vFPType, Nothing, Nothing)
          If vAP.HasControls Then
            Dim vDialogResult As DialogResult = vAP.ShowDialog()
            Dim vReportList As ParameterList = Nothing
            If vDialogResult = System.Windows.Forms.DialogResult.OK Then
              vReportList = vAP.ReturnList
              Dim vReportWhichDate As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_report_which_date, "")
              If vReportWhichDate.Length = 0 Then vReportWhichDate = "posted"
              vReportList("RP7") = vReportWhichDate
              If vReportList.Contains("StartDate") Then vReportList("RP1") = vReportList("StartDate")
              If vReportList.Contains("EndDate") Then vReportList("RP2") = vReportList("EndDate")
              If vReportList.Contains("SalesGroup") Then vReportList("RP3") = vReportList("SalesGroup")
              If vReportList.Contains("Company") Then vReportList("RP4") = vReportList("Company")
              If vReportList.Contains("Checkbox") Then
                vReportList("RP6") = vReportList("Checkbox")
                If vReportList("RP6") = "N" Then vReportList("RP6") = ""
              End If
              vReportList("ReportCode") = vReportCode
              Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
            End If
          End If

        Case CommandIndexes.mnuFinDBRepAllocatedDonations, CommandIndexes.mnuFinDBRepDonorDetails
          'These reports have parameters
          Dim vReportCode As String = ""
          Dim vReportDestination As String = ""
          Select Case pCommand
            Case CommandIndexes.mnuFinDBRepAllocatedDonations
              vReportCode = "DBAB"
            Case CommandIndexes.mnuFinDBRepDonorDetails
              vReportCode = "DBDD"
          End Select
          If vReportCode.Length > 0 Then
            Dim vAP As New frmApplicationParameters(0, vReportCode, Nothing, CareNetServices.FunctionParameterTypes.fptNone)
            If vAP.HasControls Then
              Dim vDialogResult As DialogResult = vAP.ShowDialog()
              Dim vReportList As ParameterList = Nothing
              If vDialogResult = System.Windows.Forms.DialogResult.OK Then
                vReportList = vAP.ReportParameterList
                If vReportList.ContainsKey("ReportDestination") Then
                  Select Case vReportList("ReportDestination").ToLower
                    Case "none"
                      'User chose not to run the report
                      vDialogResult = System.Windows.Forms.DialogResult.Cancel
                    Case "preview"
                      'Nothing to do
                    Case "print"
                      'Need to set the DestinationType
                      vReportList("DestinationType") = "PrintXML"
                    Case Else
                      'For everything else, store the ReportDestination
                      vReportDestination = vReportList("ReportDestination")
                  End Select
                  vReportList.Remove("ReportDestination")
                End If
              End If
              If vDialogResult = System.Windows.Forms.DialogResult.OK Then
                If pCommand = CommandIndexes.mnuFinDBRepDonorDetails Then
                  'For this report, we need to re-name the 2 parameters
                  If vReportList.ContainsKey("OrganisationNumber") Then
                    vReportList("RP1") = vReportList("OrganisationNumber")
                    vReportList.Remove("OrganisationNumber")
                  End If
                  If vReportList.ContainsKey("Relationship") Then
                    vReportList("RP2") = vReportList("Relationship")
                    vReportList.Remove("Relationship")
                  End If
                End If
                vReportList("ReportCode") = vReportCode
                Call (New PrintHandler).PrintReport(vReportList, vReportDestination)
              End If
            End If
          End If

        Case CommandIndexes.mnuFinFastDataEntryMaint
          Dim vFrmDL As New frmDisplayList(frmDisplayList.ListUsages.FastDataEntry)
          vFrmDL.ShowDialog()

        Case CommandIndexes.mnuAdminDataImport
          InitDataImport()
        Case CommandIndexes.mnuAdminMaintenanceSetup
          Dim vFrmDL As New frmDisplayList(CDBNETCL.frmDisplayList.ListUsages.MaintenanceSetup)
          vFrmDL.ShowDialog()
        Case CommandIndexes.mnuAdminMoveExternalDocuments
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftExternalDocuments, True)
        Case CommandIndexes.mnuAdminLicenceMaintenance
          Dim vForm As New frmLicenceMaintenance
          vForm.ShowDialog()
        Case CommandIndexes.mnuAdminOwnershipMaintenance
          Dim vForm As New frmOwnershipMaintenance
          vForm.ShowDialog()
        Case CommandIndexes.mnuAdminRegenerateMessageQueue
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtRegenerateMessageQueue, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminReportMaintenance
          Dim vForm As New frmReportMaintenance
          vForm.Show()
        Case CommandIndexes.mnuAdminTraderApplicationMaintenance
          Dim vForm As New frmDisplayList(CDBNETCL.frmDisplayList.ListUsages.TraderMaintenance)
          vForm.ShowDialog()
        Case CommandIndexes.mnuAdminConfigurationMaintenance
          Dim vForm As New frmConfigurationMaintenance
          vForm.ShowDialog()
        Case CommandIndexes.mnuAdminDatabaseUpgrade
          Dim vDefault As New ParameterList
          vDefault("ShowTaskStatus") = "Y"
          vDefault("DBUpgrade") = "Y"
          vDefault("BankAccount") = AppValues.DefaultTraderBankAccountCode
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtDatabaseUpgrade, vDefault, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminUpdateCustomForms
          If ShowQuestion(QuestionMessages.QmUpdateCustomFormData, MessageBoxButtons.YesNo) = DialogResult.Yes Then InitCustomFormInfo()
        Case CommandIndexes.mnuAdminUpdateGovernmentRegions
          If ShowQuestion(QuestionMessages.QmUpdateGovernmentRegion, MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Dim vDefaultParams As New ParameterList(True)
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUpdateGovernmentRegions, vDefaultParams, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
          End If
        Case CommandIndexes.mnuAdminUpdateMailsortData
          If ShowQuestion(QuestionMessages.QmUpdateMailSortData, MessageBoxButtons.YesNo) = DialogResult.Yes Then InitMailSortData()
        Case CommandIndexes.mnuAdminUpdatePaymentSchedule
          'NYI()
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtUpdatePaymentSchedule, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminUpdateTraderApplications
          If ShowQuestion(QuestionMessages.QmupdateTraderApplication, MessageBoxButtons.YesNo) = DialogResult.Yes Then CheckTraderData()
        Case CommandIndexes.mnuAdminDataUpdates
          Dim vReturnList As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptLoadDataUpdates)
          If vReturnList IsNot Nothing AndAlso vReturnList.ContainsKey("InputFilename") Then
            If DataHelper.LoadDataUpdatesFile(vReturnList).ContainsKey("Result") Then
              'File loaded OK so display Updates to user
              Dim vForm As New frmCardMaintenance(CareNetServices.XMLMaintenanceControlTypes.xmctDataUpdates)
              If vForm.ShowDialog = DialogResult.OK Then
                'User has selected some DataUpdates so run the task
                vReturnList = vForm.ReturnList
                FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtDataUpdates, vReturnList, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule)
              End If
            End If
          End If

        Case CommandIndexes.mnuAdminImportTraderApplication
          InitTraderAppInfo()
        Case CommandIndexes.mnuAdminExportCustomForm
          Dim vFrmExport As New frmExport(frmExport.ExportType.etCustomForm)
          vFrmExport.ShowDialog()
        Case CommandIndexes.mnuAdminExportReport
          Dim vFrmExport As New frmExport(frmExport.ExportType.etReports)
          vFrmExport.ShowDialog()
        Case CommandIndexes.mnuAdminExportTraderApplication
          Dim vFrmExport As New frmExport(frmExport.ExportType.etTraderApp)
          vFrmExport.ShowDialog()
        Case CommandIndexes.mnuAdminConfigurationReport
          Dim vReportList As New ParameterList(True)
          vReportList("ReportCode") = "CSANAL"
          Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case CommandIndexes.mnuAdminCheckSetup
          Dim vFrmCheckSetup As New frmCheckSetup
          vFrmCheckSetup.ShowDialog()
        Case CommandIndexes.mnuAdminCheckPaymentPlans
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCheckPaymentPlans, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuAdminPostcodeUpdate
          NYI()
        Case CommandIndexes.mnuAdminAccessControl
          Dim vForm As New frmAccessControl(ACCESS_CONTROL_VERSION)
          vForm.ShowDialog()
        Case CommandIndexes.cbiCascade
          MDIForm.LayoutMdi(MdiLayout.Cascade)
        Case CommandIndexes.cbiTileHorizontally
          MDIForm.LayoutMdi(MdiLayout.TileHorizontal)
        Case CommandIndexes.cbiTileVertically
          MDIForm.LayoutMdi(MdiLayout.TileVertical)
        Case CommandIndexes.cbiArrangeIcons
          MDIForm.LayoutMdi(MdiLayout.ArrangeIcons)
        Case CommandIndexes.cbiCloseAll
          MainHelper.CloseAll()
        Case CommandIndexes.mnuFinDBCreateUnallocatedBoxes
          Dim vList As ParameterList = Nothing
          GetDistBoxesReportData(pCommand, Nothing, Nothing, vList, CareServices.TaskJobTypes.tjtPrintBoxLabels)
          If Not vList.Count = 0 Then
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPrintBoxLabels, vList, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun)
            If vList.ValueIfSet("Print") = "Y" Then
              Dim vDistBoxesParams As New ParameterList(True)
              vDistBoxesParams("BoxNumbers") = vList("StartBoxNumber") & "," & vList("EndBoxNumber")
              vDistBoxesParams("DistributionAffiliate") = vList("DistributionAffiliate")
              Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtDistributionBoxes, vDistBoxesParams)
              vList("StartBoxNumber") = vDT.Rows(0)("DistributionCode").ToString
              vList("EndBoxNumber") = vDT.Rows(vDT.Rows.Count - 1)("DistributionCode").ToString
              vList("ShowResultMessage") = "N"
              RunDistributionBoxesReport(CommandIndexes.mnuFinDBPrintBoxLabels, vList)
            End If
          End If
        Case CommandIndexes.mnuFinDBSetShippingInformation
          Dim vList As ParameterList = Nothing
          GetDistBoxesReportData(pCommand, Nothing, Nothing, vList, CareServices.TaskJobTypes.tjtShipDistributionBoxes)
          If Not vList.Count = 0 Then
            If vList.Contains("BoxDestination") Then
              vList("BoxDestination") = """" & vList("BoxDestination") & """"
            End If
            If vList.Contains("Comments") Then
              vList("Comments") = """" & vList("Comments") & """"
            End If
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtShipDistributionBoxes, vList, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun)
            If vList.ValueIfSet("Print") = "Y" Then
              RunDistributionBoxesReport(CommandIndexes.mnuFinDBSetShippingInformation, vList)
            End If
          End If
        Case CommandIndexes.cbiHelp
          Dim vForm As New frmBrowser(DataHelper.HelpURL("index.htm", True), False, True)
          vForm.Show()
        Case CommandIndexes.cbiReleaseNotes
          Dim vForm As New frmBrowser(DataHelper.HelpURL("release_notes.htm", True), False, True)
          vForm.Show()
        Case CommandIndexes.cbiKnowledgebase
          Dim vForm As New frmBrowser(My.Settings.KnowledgebaseUrl, False, True)
          vForm.Show()
        Case CommandIndexes.cbiSupportForum
          Dim vForm As New frmBrowser(My.Settings.SupportForumUrl, False, True)
          vForm.Show()
        Case CommandIndexes.cbiAbout
          Dim vForm As New frmAboutBox
          vForm.ShowDialog(mvParentForm)
        Case CommandIndexes.mnuFinDBPrintThankYouLetters, CommandIndexes.mnuFinDBPrintAdviceNotes, CommandIndexes.mnuFinDBPrintBoxLabels,
             CommandIndexes.mnuFinDBPrintPackingSlips
          RunDistributionBoxesReport(pCommand)
        Case CommandIndexes.mnuFinDBSetArrivalInformation
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtSetBoxesArrived, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.cbiTableMaintenance
          Dim vForm As New frmTableMaintenance()
          vForm.Show()
        Case CommandIndexes.cbiJobSchedule
          Dim vForm As New frmJobProcessor
          vForm.ShowDialog()
        Case CommandIndexes.cbiPostcodeProximity
          Dim vForm As New frmFinder(CType(CareNetServices.XMLDataFinderTypes.xdftPostcodeProximity, CareServices.XMLDataFinderTypes), Nothing)
          vForm.ShowDialog()
        Case CommandIndexes.mnuEventBlockBooking
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CType(CareNetServices.TaskJobTypes.tjtEventBlockBooking, CareServices.TaskJobTypes), vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case CommandIndexes.mnuExamsMaintenance
          FormHelper.ShowExamIndex()
        Case CommandIndexes.mnuExamEnterResults
          Dim vForm As New frmExamResults
          vForm.ShowDialog()
        Case CommandIndexes.mnuExamAllocateCandidateNumbers
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtExamAllocateCandidateNumbers, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamAllocateMarkers
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtExamAllocateMarkers, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamApplyGrading
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtExamApplyGrading, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamGenerateExemptionInvoices
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtExamGenerateExemptionInvoices, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamLoadCSVResults
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtExamLoadCsvResults, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamCancelProvisionalBookings
          Dim vDefaults As New ParameterList(True)
          vDefaults("ExamBooking") = "Y"
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtCancelProvisionalBookings, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamProcessCertificates
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtProcessCertificateData, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuExamGenerateCertificates
          Call New GeneralMailing(CareNetServices.MailingTypes.mtyExamCertificates, CareServices.TaskJobTypes.tjtExamCertificates).Process(0)
        Case CommandIndexes.mnuExamSheduleCertificateReprints
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtExamCertificateReprints, vDefaults, False, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)

        Case CommandIndexes.mnuEventCancel
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtCancelEvent, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuCpdApplyPoints
          Dim vDefaults As New ParameterList(True)
          vDefaults("ActivityCPD") = "Y"
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtApplyCPDPoints, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.mnuEventCancelProvisionalTransaction
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtCancelProvisionalBookings, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case CommandIndexes.cbiRunTest
          RunTest()

        Case CommandIndexes.mnuInternalCheckNonCoreTables
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtCheckNonCoreTables, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case CommandIndexes.mnuInternalGenerateTableCreationFiles
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGenerateTableCreationFiles, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case CommandIndexes.mnuInternalGetReportData
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGetReportData, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
        Case CommandIndexes.mnuInternalGetConfigNameData
          Dim vDefaults As New ParameterList(True)
          FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtGetConfigNameData, vDefaults, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
      End Select
    Catch vException As Exception
      DataHelper.HandleJobException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub vForm_SelectedContactChanged(ByVal sender As Object, ByVal pContactNumber As Integer)
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub
  Private Sub NYI()
    ShowInformationMessage(InformationMessages.ImNotYetImplemented)
  End Sub

  Private Sub ProcessExplorerItemFromMenu(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vItem As ToolStripItem = DirectCast(sender, ToolStripItem)
    Dim vTag As String = vItem.Tag.ToString
    Dim vURl As String = vTag.Substring(0, vTag.Length - 2)
    Dim vShowToolbar As Boolean = vTag.EndsWith(",Y")

    ProcessExplorerItem(vURl, vShowToolbar)

  End Sub

  Private Sub ProcessExplorerItemFromCommand(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vItem As MenuToolbarCommand = DirectCast(sender, MenuToolbarCommand)
    Dim vTag As String = vItem.CommandName
    Dim vURl As String = vTag.Substring(0, vTag.Length - 2)
    Dim vShowToolbar As Boolean = vURl.EndsWith(",Y")

    ProcessExplorerItem(vURl, vShowToolbar)

  End Sub
  Private Sub ProcessExplorerItem(pURl As String, pShowToolbar As Boolean)
    Try
      Dim vForm As New frmBrowser(pURl, pShowToolbar)
      vForm.Show()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub RunDistributionBoxesReport(ByVal pReportCommand As CommandIndexes, Optional ByVal pList As ParameterList = Nothing)
    Dim vTypeList As ParameterList = Nothing
    Dim vReportList As ParameterList = Nothing

    GetDistBoxesReportData(pReportCommand, vTypeList, vReportList, pList)

    If pList IsNot Nothing AndAlso pList.Count > 0 Then
      With vReportList
        .AddItemIfValueSet("RP1", pList.ValueIfSet("DistributionAffiliate"))
        .AddItemIfValueSet("RP2", pList.ValueIfSet("StartBoxNumber"))
        .AddItemIfValueSet("RP3", pList.ValueIfSet("EndBoxNumber"))
        .AddItemIfValueSet("RP4", pList.ValueIfSet("Print"))
        .AddItemIfValueSet("RP5", pList.ValueIfSet("Reprint"))
        .AddItemIfValueSet("RP6", pList.ValueIfSet("FromProcessedDate"))
        .AddItemIfValueSet("RP7", pList.ValueIfSet("ToProcessedDate"))
        .AddItemIfValueSet("RP8", pList.ValueIfSet("PrintedOn"))
        .AddItemIfValueSet("RP9", IIf(vTypeList.ValueIfSet("DistBoxReportType") = "0", "Y", "").ToString)
        .AddItemIfValueSet("RP10", IIf(vTypeList.ValueIfSet("DistBoxReportType") = "3", "Y", "").ToString)
        .AddItemIfValueSet("RP11", pList.ValueIfSet("FromBatchNumber"))
        .AddItemIfValueSet("RP12", pList.ValueIfSet("ToBatchNumber"))
        .AddItemIfValueSet("RP13", pList.ValueIfSet("Mailing"))
      End With

      Dim vFileName As String
      If BooleanValue(pList.ValueIfSet("CSVOutput")) Then
        vFileName = pList("ReportDestination")
      Else
        vFileName = DataHelper.GetTempFile(".csv")
      End If
      DataHelper.GetReportFile(vReportList, vFileName)

      Dim vFileInfo As New FileInfo(vFileName)
      Dim vReader As New FileReader(vFileName)
      Dim vHeader As String = vReader.ReadLine()
      Dim vProcessJob As Boolean

      If BooleanValue(pList.ValueIfSet("CSVOutput")) Then
        vProcessJob = Not vReader.EndOfFile
        vReader.CloseFile()
      Else
        Dim vMax As Integer = IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.opt_max_document_count, MAX_DOCUMENT_COUNT.ToString))
        While vReader.EndOfFile = False
          Dim vTempFile As String = DataHelper.GetTempFile(".csv")
          Dim vWriter As New StreamWriter(vTempFile, False, Encoding.Default) 'Use Encoding.Default to read the Accents correctly
          vWriter.WriteLine(vHeader)
          For vIndex As Integer = 1 To vMax
            If vReader.EndOfFile = True Then
              Exit For
            End If
            vWriter.WriteLine(vReader.ReadLine)
          Next
          vWriter.Close()
          Dim vApplication As ExternalApplication = GetDocumentApplication(pList.ValueIfSet("DocFileExtension"))
          vApplication.MergeStandardDocument(pList("StandardDocument"), pList.ValueIfSet("DocFileExtension"), vTempFile, True)
          vProcessJob = True
        End While
        vReader.CloseFile()
        vFileInfo.Delete()
      End If


      If vProcessJob Then
        pList.Add("DistBoxReportType", vTypeList("DistBoxReportType"))
        If pList.ValueIfSet("PrintedOn").Length > 0 Then pList("PrintedOn") = """" & pList.ValueIfSet("PrintedOn") & """"
        FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtDistributionBoxReports, pList, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun, True)
      ElseIf Not pList.Contains("ShowResultMessage") Then
        ShowInformationMessage(InformationMessages.ImNoRecordsMatch)
      End If
    End If
  End Sub

  Private Shared Sub GetDistBoxesReportData(ByVal pReportCommand As CommandIndexes, ByRef pTypeList As ParameterList, ByRef pReportList As ParameterList, ByRef pList As ParameterList, Optional ByVal pTaskJobType As CareServices.TaskJobTypes = CareServices.TaskJobTypes.tjtDistributionBoxReports)
    pTypeList = New ParameterList(True)
    pReportList = New ParameterList(True)
    pTypeList.Add("DistBoxReportType", "")
    pTypeList.Add("MailmergeHeader", "")
    pReportList.Add("ReportCode", "")
    Select Case pReportCommand
      Case CommandIndexes.mnuFinDBPrintThankYouLetters, CommandIndexes.mnuFinDBPrintAdviceNotes, CommandIndexes.mnuFinDBSetShippingInformation
        pReportList("ReportCode") = "DBAN"
        pTypeList("MailmergeHeader") = "DBTYAN"
        If pReportCommand = CommandIndexes.mnuFinDBPrintThankYouLetters Then
          pTypeList("DistBoxReportType") = "3"
        Else
          pTypeList("DistBoxReportType") = "0"
        End If
      Case CommandIndexes.mnuFinDBPrintBoxLabels, CommandIndexes.mnuFinDBCreateUnallocatedBoxes
        pReportList("ReportCode") = "DBPL"
        pTypeList("MailmergeHeader") = "DBBLBL"
        pTypeList("DistBoxReportType") = "1"
      Case CommandIndexes.mnuFinDBPrintPackingSlips
        pReportList("ReportCode") = "DBPS"
        pTypeList("MailmergeHeader") = "DBPSLP"
        pTypeList("DistBoxReportType") = "2"
    End Select


    If pList Is Nothing Then
      'BR13332: Check any custom mailmerge header
      Dim vReports As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtReports, pReportList)
      If vReports IsNot Nothing AndAlso vReports.Rows.Count > 1 Then
        vReports.DefaultView.RowFilter = "ClientCode <> ''"
        vReports = vReports.DefaultView.ToTable
        If vReports.Rows.Count > 0 AndAlso vReports.Rows(0)("MailmergeHeader").ToString.Length > 0 Then
          pTypeList("MailmergeHeader") = vReports.Rows(0)("MailmergeHeader").ToString
        End If
      End If
      pList = FormHelper.ShowApplicationParameters(pTaskJobType, pTypeList)
    End If
  End Sub

  Private Sub LaunchTraderApplicationFromMenu(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim vItem As ToolStripItem = DirectCast(sender, ToolStripItem)
    Dim vTraderApplicationNumber As Integer = IntegerValue(vItem.Tag.ToString)

    LaunchTraderApplication(vTraderApplicationNumber)
  End Sub

  Private Sub LaunchTraderApplicationFromCommand(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vItem As MenuToolbarCommand = DirectCast(sender, MenuToolbarCommand)
    Dim vTraderApplicationNumber As Integer = vItem.CommandID

    LaunchTraderApplication(vTraderApplicationNumber)
  End Sub

  Private Sub LaunchTraderApplication(pTraderApplicationNumber As Integer)
    Dim vCursor As New BusyCursor()
    Try
      ProcessTraderApplication(pTraderApplicationNumber)
    Catch vCareException As CareException
      If vCareException.ErrorNumber = CareException.ErrorNumbers.enTraderUnsupportedFeature Or vCareException.ErrorNumber = CareException.ErrorNumbers.enSpecificTraderApplicationUnsupported Or
        vCareException.ErrorNumber = CareException.ErrorNumbers.enCreditStatementGenerationUnsupported Or vCareException.ErrorNumber = CareException.ErrorNumbers.enInvalidConfig Then
        ShowInformationMessage(vCareException.Message)
      Else
        DataHelper.HandleException(vCareException)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub ProcessTraderApplication(pTraderApplicationNumber As Integer)

    'Initialise the trader application and check if it is batch led
    Dim vTraderApplication As New TraderApplication(pTraderApplicationNumber)
    Dim vBatchNumber As Integer
    Dim vNewBatchCreated As Boolean
    If vTraderApplication.BatchLedApp Then
      'It is batch led but we don't yet have a batch number so go and find one
      Dim vList As New ParameterList
      vList.IntegerValue("TraderApplication") = pTraderApplicationNumber
      vList("BatchType") = vTraderApplication.BatchTypeCode
      If Not String.IsNullOrEmpty(vTraderApplication.BatchPaymentMethod) Then
        vList("PaymentMethod") = vTraderApplication.BatchPaymentMethod
      End If
      If FormHelper.ShowBatchFinder(vBatchNumber, vList, mvParentForm) Then
        'If we decided to create a new batch then default a lot of values from the trader application and create a new batch
        If vBatchNumber = 0 Then
          vBatchNumber = FormHelper.CreateNewTraderBatch(vTraderApplication, mvParentForm)
          vNewBatchCreated = (vBatchNumber > 0)
        End If
      End If
      If vBatchNumber > 0 Then
        'Got a batch number now so get the transactions from the batch
        vTraderApplication = New TraderApplication(pTraderApplicationNumber, vBatchNumber)
        Dim vBatchInfo As BatchInfo = New BatchInfo(vBatchNumber, vNewBatchCreated)
        Dim vForm As New frmTraderTransactions(vTraderApplication, vBatchInfo)
        'Show the transactions form and await the user
        EnableTraderApplications(False)
        vForm.Show()
        If vNewBatchCreated Then vForm.RunTraderForNewBatch()
      End If
    ElseIf vTraderApplication.ApplicationType = ApplicationTypes.atChequeProcessing Then
      FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtChequeProduction, Nothing, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
    Else
      EnableTraderApplications(False)
      FormHelper.RunTraderApplication(vTraderApplication)
    End If
  End Sub

  Private Sub ProcessFastDataEntryFromMenu(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim vItem As ToolStripItem = DirectCast(sender, ToolStripItem)
    Dim vFDEPageNumber As Integer = IntegerValue(vItem.Tag.ToString)

    ProcessFastDataEntry(vFDEPageNumber)

  End Sub
  Private Sub ProcessFastDataEntryFromCommand(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim vItem As MenuToolbarCommand = DirectCast(sender, MenuToolbarCommand)
    Dim vFDEPageNumber As Integer = vItem.CommandID

    ProcessFastDataEntry(vFDEPageNumber)

  End Sub
  Private Sub ProcessFastDataEntry(pFDEPageNumber As Integer)
    Dim vCursor As New BusyCursor
    Try
      EnableTraderApplications(False)
      FormHelper.RunFastDataEntry(pFDEPageNumber)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub CopyEventPricingMatrix()
    Dim vForm As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptCopyEventPricingMatrix, New ParameterList, New ParameterList)
    If vForm.ShowDialog(mvParentForm) = System.Windows.Forms.DialogResult.OK Then
      Dim vBusyCursor As New BusyCursor()
      Try
        Dim vList As ParameterList = vForm.ReturnList
        Dim vReturn As ParameterList = DataHelper.CopyEventPricingMatrix(vList)
        DataHelper.ClearCachedTable("EventPricingMatrices", CareNetServices.XMLLookupDataTypes.xldtEventPricingMatrices)
        Dim vEPM As String = ""
        If vReturn.ContainsKey("EventPricingMatrix") Then vEPM = vReturn("EventPricingMatrix")
        ShowInformationMessage(InformationMessages.ImEventPricingMatrixCreated, vEPM)
      Catch vCareEX As CareException
        Select Case vCareEX.ErrorNumber
          Case CareException.ErrorNumbers.enEventPricingMatrixAlreadyExists, CareException.ErrorNumbers.enFromAndToPricingMatrixMustBeDifferent, CareException.ErrorNumbers.enUnableToCopyEventPricingMatrix
            ShowErrorMessage(vCareEX.Message)
          Case Else
            DataHelper.HandleException(vCareEX)
        End Select
      Catch vEX As Exception
        DataHelper.HandleException(vEX)
      Finally
        vBusyCursor.Dispose()
      End Try
    End If
  End Sub

  Private Sub InitTraderAppInfo(Optional ByVal pFileName As String = "", Optional ByVal pRunType As String = "", Optional ByVal pConfirmReplace As Boolean = False)
    Dim vComplete As Boolean
    Dim vList As New ParameterList(True)
    Try
      Dim vOFD As New OpenFileDialog
      Dim vSelectFile As Boolean = True
      If pFileName.Length = 0 Then
        With vOFD
          Do While vSelectFile
            .Title = ControlText.OfdSelectInputFile
            .CheckFileExists = True
            .CheckPathExists = True
            .FileName = ""
            .Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
              If .FileName.StartsWith("\\") Then
                pFileName = .FileName
                vSelectFile = False
              Else
                'only UNC files are supported allow the user to reselect a file
                ShowInformationMessage(InformationMessages.ImUNCPathOnly)
              End If
            Else
              vSelectFile = False
            End If
          Loop
        End With
      End If

      If pFileName.Length > 0 Then
        If pRunType.Length = 0 Then
          Dim vRename As New ParameterList
          vRename("None2") = "Do you want to overwrite the existing applications?"
          vRename("RunType") = "Do Not Import Existing Trader Applications"
          vRename("RunType2") = "Overwrite Existing Trader Applications?"
          vRename("RunType3") = "Import Existing Applications with a new number"
          Dim vDefaults As New ParameterList
          vDefaults("RunType") = "D"
          vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptImportTraderApp, vDefaults, vRename)
        Else
          vList("RunType") = pRunType
        End If

        If vList.Count > 0 Then
          If pConfirmReplace Then vList("ConfirmReplaceData") = "Y"
          vList("FileName") = pFileName
          Dim vReturnList As ParameterList = DataHelper.ImportTraderAppInfo(vList)
          If vReturnList.ContainsKey("InvalidTraderApps") Then
            ShowInformationMessage(InformationMessages.ImTraderAppImportFail)
            For Each vTraderAppNo As String In vReturnList("InvalidTraderApps").ToString.Split(","c)
              Dim vForm As New frmFPApplication(IntegerValue(vTraderAppNo), True)
              vForm.ShowDialog()
            Next
          End If
          vComplete = True
        End If
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enReplaceData
          If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = DialogResult.Yes Then
            InitTraderAppInfo(vList("FileName"), vList("RunType"), True)
          End If
        Case CareException.ErrorNumbers.enTraderAppReadOnly
          ShowInformationMessage(vEx.Message)
          vComplete = True
        Case Else
          DataHelper.HandleException(vEx)
      End Select
    Finally
      If vComplete Then ShowInformationMessage(InformationMessages.ImTraderAppImportComplete)
    End Try
  End Sub

  Private Sub InitMailSortData()
    Dim vList As New ParameterList(True)
    Dim vOFD As New OpenFileDialog
    Dim vSelectFile As Boolean = True
    With vOFD
      Do While vSelectFile
        .Title = ControlText.OfdSelectInputFile
        .CheckFileExists = True
        .CheckPathExists = True
        .FileName = ""
        .Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv"
        If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
          If .FileName.StartsWith("\\") Then
            vList("FileName") = .FileName
            DataHelper.InitMailsortData(vList)
            ShowInformationMessage(InformationMessages.ImMailSortDataImportComplete)
            vSelectFile = False
          Else
            'only UNC files are supported allow the user to reselect a file
            ShowInformationMessage(InformationMessages.ImUNCPathOnly)
          End If
        Else
          vSelectFile = False
        End If
      Loop
    End With
  End Sub

  Private Sub InitCustomFormInfo(Optional ByVal pFileName As String = "", Optional ByVal pConfirmReplace As Boolean = False)
    Dim vList As New ParameterList(True)
    Try
      Dim vOFD As New OpenFileDialog
      Dim vSelectFile As Boolean = True
      If pFileName.Length = 0 Then
        With vOFD
          Do While vSelectFile
            .Title = ControlText.OfdSelectInputFile
            .CheckFileExists = True
            .CheckPathExists = True
            .FileName = ""
            .Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
              If .FileName.StartsWith("\\") Then
                pFileName = .FileName
                vSelectFile = False
              Else
                'only UNC files are supported allow the user to reselect a file
                ShowInformationMessage(InformationMessages.ImUNCPathOnly)
              End If
            Else
              vSelectFile = False
            End If
          Loop
        End With
      End If

      If pFileName.Length > 0 Then
        If pConfirmReplace Then vList("ConfirmReplaceData") = "Y"
        vList("FileName") = pFileName
        DataHelper.InitCustomForms(vList)
        ShowInformationMessage(InformationMessages.ImCustomFormImportComplete)
      End If
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enReplaceData Then
        If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = DialogResult.Yes Then
          InitCustomFormInfo(vList("FileName"), True)
        End If
      Else
        Throw vEx
      End If
    End Try
  End Sub

  ''' <summary>Runs report for Reprint Claim details, Claim Analysis, Sponsorship Event details  and Irish Gift Aid Claim details</summary>
  ''' <param name="pReportList"></param>
  ''' <remarks></remarks>
  Private Sub RunClaimReport(ByRef pReportList As ParameterList)
    Try
      Dim vCaption As String = String.Empty
      Select Case pReportList("ReportCode").ToString
        Case "GACANA"
          vCaption = ControlText.GiftAidClaimAnalysisCaption
        Case "GASP"
          vCaption = ControlText.GiftAidReprintSponsoredEventCaption
        Case "IGAT"
          vCaption = ControlText.GiftAidReprintIrishTaxClaim
        Case Else
          vCaption = ControlText.GiftAidReprintTaxClaim
      End Select

      Dim vDefaults As New ParameterList()
      vDefaults("ReportCode") = pReportList("ReportCode").ToString
      If Not String.IsNullOrWhiteSpace(vCaption) Then vDefaults("Caption") = vCaption
      If pReportList.ContainsKey("RPpAdjustmentText") Then vDefaults("AdjustmentText") = pReportList("RPpAdjustmentText").ToString

      Dim vResultList As ParameterList = FormHelper.ShowApplicationParameters(CareNetServices.FunctionParameterTypes.fptGAReprintTaxClaim, vDefaults)
      If vResultList IsNot Nothing AndAlso vResultList.ContainsKey("ClaimNumber") Then
        Dim vClaimNumber As Integer = IntegerValue(vResultList("ClaimNumber").ToString)
        If pReportList.ContainsKey("RPpAdjustmentText") AndAlso vResultList.ContainsKey("AdjustmentText") Then pReportList("RPpAdjustmentText") = vResultList("AdjustmentText")
        If vClaimNumber > 0 Then
          pReportList("RP1") = vClaimNumber.ToString
          Call (New PrintHandler).PrintReport(pReportList, PrintHandler.PrintReportOutputOptions.AllowSave)

          ShowInformationMessage(InformationMessages.ImGiftAidReportCompleted)
        End If
      End If

    Catch vEx As OverflowException
      'This exception comes when user try to enter an value which is not crossing the max integer value limit
      ShowInformationMessage(InformationMessages.ImInvalidClaimNumber)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub CheckTraderData()
    Try
      Dim vUpdate As Boolean = False
      Dim vAppVersion As Integer = AppValues.CurrentTraderApplicationVersion()
      If AppValues.SystemTraderApplicationVersion > vAppVersion Then
        vUpdate = True
      Else
        If ShowQuestion(QuestionMessages.QmConfirmUpdateTraderApplication, MessageBoxButtons.YesNo) = DialogResult.Yes Then
          vUpdate = True
        End If
      End If

      Dim vList As New ParameterList(True)
      If vUpdate Then
        Dim vDefaultBankAccount As String = AppValues.DefaultTraderBankAccountCode
        If vDefaultBankAccount.Length > 0 Then
          vList("BankAccount") = vDefaultBankAccount
          vUpdate = True
        Else
          Dim vSI As New frmSimpleFinder
          vSI.Init(CareNetServices.XMLLookupDataTypes.xldtBankAccounts, False)
          If vSI.ShowDialog() = DialogResult.OK Then
            vList("BankAccount") = vSI.ResultValue
            vUpdate = True
          Else
            vUpdate = False
          End If
        End If
      End If

      If vUpdate Then
        vList.IntegerValue("Version") = vAppVersion
        vList.IntegerValue("NewVersion") = AppValues.SystemTraderApplicationVersion
        Dim vDatatable As DataTable = DataHelper.CheckTraderData(vList)
        If vDatatable IsNot Nothing Then ShowInformationMessage(InformationMessages.ImUpdateTraderApplication)
      End If

    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub RunBatchReport(ByVal pReportCode As String)
    Dim vResult As String = String.Empty
    Dim vReportList As New ParameterList(True)

    'Show batch finder to 
    Dim vBatchNumber As Integer
    Dim vFinderList As New ParameterList
    vFinderList("AllowNew") = "N"
    vFinderList("SuppressBatchView") = "Y"
    FormHelper.ShowBatchFinder(vBatchNumber, vFinderList, mvParentForm)
    If vBatchNumber > 0 Then
      vReportList("RPbatch_number") = vBatchNumber.ToString
      vReportList("ReportCode") = pReportCode
      'Run the report 
      Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
    End If

  End Sub

  Private Sub InitDataImport()
    Dim vParams As New ParameterList(True)
    Dim vInitialFolder As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.default_import_directory, String.Empty)
    Dim vOFD As New OpenFileDialog
    Dim vSelectFile As Boolean = True

    With vOFD
      Do While vSelectFile
        .Title = ControlText.OfdSelectImportFile
        .CheckFileExists = True
        .CheckPathExists = True
        .FileName = ""
        .Filter = "Import Files(*.csv;*.fff;*.def)|*.csv;*.fff;*.def|CSV Files (*.csv)|*.csv|Fixed Format Files(*.fff)|*.fff|Definition Files(*.def)|*.def|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 1
        If Not String.IsNullOrWhiteSpace(vInitialFolder) Then
          .InitialDirectory = vInitialFolder
          vInitialFolder = String.Empty
        End If
        If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
          If vOFD.FileName = String.Empty Then
            ShowWarningMessage(InformationMessages.ImImportFileNotFound)
          ElseIf Not .FileName.StartsWith("\\") Then
            'only UNC files are supported allow the user to reselect a file
            ShowInformationMessage(InformationMessages.ImUNCPathOnly)
          Else
            Dim vFrmImport As New frmImport(vOFD.FileName)
            vFrmImport.Show()
            vSelectFile = False
          End If
        Else
          vSelectFile = False
        End If
      Loop
    End With
  End Sub

#End Region

#Region "Public Properties"

  Public ReadOnly Property SeparatorIndex() As Integer
    Get
      Return CommandIndexes.cbiSeparator
    End Get
  End Property

  Public Sub EnableTraderApplications(ByVal pEnable As Boolean)
    If Not pEnable AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_multiple_trader_application) Then
      pEnable = True
    End If
    mvApplicationsMenu.Enabled = pEnable
    If FormView = FormViews.Modern Then
      Dim vTraderApps As IEnumerable(Of MenuToolbarCommand) = mvMenuItems.Cast(Of MenuToolbarCommand).Where(Function(command) command.ExplorerMenuAttribute IsNot Nothing AndAlso command.ExplorerMenuAttribute.Section = ExplorerMenuSection.Trader)
      If vTraderApps IsNot Nothing Then
        For Each vTraderApp As MenuToolbarCommand In vTraderApps
          vTraderApp.IsEnabled = pEnable
        Next
      End If
    End If
  End Sub

  Public WriteOnly Property DashboardChecked() As Boolean
    Set(ByVal value As Boolean)
      mvMenuItems(CommandIndexes.cbiDashboard.ToString).CheckToolStripItem(mvMenuStrip, Nothing, value)
    End Set
  End Property
  Public WriteOnly Property HeaderPanelChecked() As Boolean
    Set(ByVal value As Boolean)
      mvMenuItems(CommandIndexes.cbiHeaderPanel.ToString).CheckToolStripItem(mvMenuStrip, Nothing, value)
    End Set
  End Property
  Public WriteOnly Property SelectionPanelChecked() As Boolean
    Set(ByVal value As Boolean)
      mvMenuItems(CommandIndexes.cbiSelectionPanel.ToString).CheckToolStripItem(mvMenuStrip, Nothing, value)
    End Set
  End Property
  Public WriteOnly Property StatusBarChecked() As Boolean
    Set(ByVal value As Boolean)
      mvMenuItems(CommandIndexes.cbiStatusBar.ToString).CheckToolStripItem(mvMenuStrip, Nothing, value)
    End Set
  End Property
  Public WriteOnly Property ToolBarChecked() As Boolean
    Set(ByVal value As Boolean)
      mvToolStrip.Visible = value
      mvMenuItems(CommandIndexes.cbiToolbar.ToString).CheckToolStripItem(mvMenuStrip, Nothing, value)
    End Set
  End Property
  Public WriteOnly Property NavigationPanelChecked() As Boolean
    Set(ByVal value As Boolean)
      mvMenuItems(CommandIndexes.cbiNavigationPanel.ToString).CheckToolStripItem(mvMenuStrip, Nothing, value)
    End Set
  End Property
  Public ReadOnly Property MenuItems As CollectionList(Of MenuToolbarCommand)
    Get
      Return mvMenuItems
    End Get
  End Property
#End Region

#Region "Toolstrip"

  Public Sub BuildToolbar()
    If Settings.MainToolbarTextPosition <> 0 Then mvToolbarTextPosition = CType(Settings.MainToolbarTextPosition, TextImageRelation)
    Dim vToolbarSet As Boolean
    Try
      If Settings.MainToolbarItems.Length > 0 Then
        Dim vCustomItems() As String = Settings.MainToolbarItems.Split(","c)
        Dim vHasTipText As Boolean = Settings.MainToolbarItemsTipText.Length > 0
        Dim vItemToolTips() As String = Settings.MainToolbarItemsTipText.Split(","c)
        Dim vHasLabelText As Boolean = Settings.MainToolbarItemsText.Length > 0
        Dim vItemToolText() As String = Settings.MainToolbarItemsText.Split(","c)
        Dim vIndex As Integer
        For Each vItem As String In vCustomItems
          If vHasTipText Then
            mvMenuItems([Enum].GetName(GetType(CommandIndexes), CInt(vItem))).ToolTipText = vItemToolTips(vIndex)
          End If
          If vHasLabelText Then
            mvMenuItems([Enum].GetName(GetType(CommandIndexes), CInt(vItem))).ToolBarText = vItemToolText(vIndex)
          End If
          vIndex += 1
          mvMenuItems([Enum].GetName(GetType(CommandIndexes), CInt(vItem))).AddToToolStrip(mvToolStrip)
        Next
        vToolbarSet = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
    If Not vToolbarSet Then
      Dim vItems As New List(Of String)
      Dim vThamesDefaultIcons() As String = {
          CommandIndexes.cbiDashboard.ToString,
          CommandIndexes.cbiContactFinder.ToString,
          CommandIndexes.cbiOrganisationFinder.ToString,
          CommandIndexes.cbiDocumentFinder.ToString,
          CommandIndexes.cbiActionFinder.ToString,
          CommandIndexes.cbiSelectionSetFinder.ToString,
          CommandIndexes.cbiEventFinder.ToString,
          CommandIndexes.cbiSeparator.ToString,
          CommandIndexes.cbiNewDocument.ToString,
          CommandIndexes.cbiNewTelephoneCall.ToString,
          CommandIndexes.cbiSeparator.ToString,
          CommandIndexes.cbiListManager.ToString}
      vItems.AddRange(vThamesDefaultIcons)

      For Each vItem As String In vItems
        mvMenuItems(vItem).SetDefaults()
        mvMenuItems(vItem).ToolBarText = mvMenuItems(vItem).ToolTipText
        mvMenuItems(vItem).AddToToolStrip(mvToolStrip)
      Next
    End If
    For Each vItem As ToolStripItem In mvToolStrip.Items
      If vItem.Text.Length > 0 Then
        vItem.AccessibleName = vItem.Text
      Else
        vItem.AccessibleName = vItem.ToolTipText
      End If
      vItem.TextImageRelation = mvToolbarTextPosition
    Next
    LargeToolbarIcons = Settings.LargeToolbarIcons
  End Sub

  Public Sub ResetToolbar()
    mvToolStrip.Items.Clear()
    BuildToolbar()
  End Sub

  Public Sub SaveToolbarItems()
    Dim vSB As New StringBuilder
    Dim vToolTipSB As New StringBuilder
    Dim vToolTextSB As New StringBuilder
    Dim vCommand As MenuToolbarCommand
    For Each vItem As ToolStripItem In mvToolStrip.Items
      If vSB.Length > 0 Then
        vSB.Append(",")
        vToolTipSB.Append(",")
        vToolTextSB.Append(",")
      End If
      vCommand = TryCast(vItem.Tag, MenuToolbarCommand)
      vSB.Append(vCommand.CommandID)
      vToolTipSB.Append(vCommand.ToolTipText.Replace(",", ""))
      vToolTextSB.Append(vCommand.ToolBarText.Replace(",", ""))
    Next
    Settings.MainToolbarItems = vSB.ToString
    Settings.MainToolbarItemsTipText = vToolTipSB.ToString
    Settings.MainToolbarItemsText = vToolTextSB.ToString
    Settings.MainToolbarTextPosition = MainHelper.ToolbarTextPosition
    Settings.Save()
  End Sub

  Public Property ToolbarTextPosition() As TextImageRelation
    Get
      Return mvToolbarTextPosition
    End Get
    Set(ByVal pValue As TextImageRelation)
      If pValue <> mvToolbarTextPosition Then
        For Each vItem As ToolStripItem In mvToolStrip.Items
          vItem.TextImageRelation = pValue
        Next
      End If
      mvToolbarTextPosition = pValue
    End Set
  End Property

  Public Property LargeToolbarIcons() As Boolean
    Get
      Return mvToolStrip.ImageList Is mvImageList32
    End Get
    Set(ByVal Value As Boolean)
      If Value Then
        mvToolStrip.ImageScalingSize = New Size(32, 32)
        mvToolStrip.ImageList = mvImageList32
      Else
        mvToolStrip.ImageScalingSize = New Size(16, 16)
        mvToolStrip.ImageList = mvImageList16
      End If
    End Set
  End Property

  Private Sub ToolbarDoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
    LargeToolbarIcons = Not LargeToolbarIcons
    Settings.LargeToolbarIcons = LargeToolbarIcons
  End Sub
  Private Sub ToolbarDragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
    If e.Data.GetDataPresent(GetType(MenuToolbarCommand).FullName) Then e.Effect = DragDropEffects.Copy
  End Sub
  Private Sub ToolbarDragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
    If e.Data.GetDataPresent(GetType(MenuToolbarCommand).FullName) Then
      Dim vCommand As MenuToolbarCommand = CType(e.Data.GetData(GetType(MenuToolbarCommand).FullName), MenuToolbarCommand)
      'See if the command that was dropped is already on the toolbar
      Dim vCheckCommand As MenuToolbarCommand
      For Each vCheckItem As ToolStripItem In mvToolStrip.Items
        vCheckCommand = TryCast(vCheckItem.Tag, MenuToolbarCommand)
        If vCheckCommand IsNot Nothing AndAlso vCheckCommand.CommandID = vCommand.CommandID AndAlso vCheckCommand.CommandID <> Me.SeparatorIndex Then
          'We found the command already exists in the toolbar so remove it
          mvToolStrip.Items.Remove(vCheckItem)
          Exit For
        End If
      Next
      Dim vItem As ToolStripItem = TryCast(mvToolStrip.GetItemAt(mvToolStrip.PointToClient(New Point(e.X, e.Y))), ToolStripItem)
      If vItem IsNot Nothing Then
        vItem = vCommand.AddToToolStripAt(mvToolStrip, mvToolStrip.Items.IndexOf(vItem))    'On top of another item so place in front
      Else
        vItem = vCommand.AddToToolStrip(mvToolStrip)                                'On space at end of the toolbar so add at end
      End If
      If vItem IsNot Nothing Then
        vItem.TextImageRelation = mvToolbarTextPosition
        If vItem.Text.Length > 0 Then
          vItem.AccessibleName = vItem.Text
        Else
          vItem.AccessibleName = vItem.ToolTipText
        End If
      End If
    End If
    SaveToolbarItems()
  End Sub

#End Region

#Region "Test Function"

  Private Sub RunTest()

  End Sub

#End Region

  Private Sub LoadWorkstreamMenu(pParentMenu As ToolStripMenuItem)

    Dim vWorkstreamGroupList As New SortedList

    'Get the data
    Dim vWorkstreamsTable As DataTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtWorkstreamGroups)
    If Not vWorkstreamsTable Is Nothing Then
      Dim vRows As DataRow() = vWorkstreamsTable.Select()
      If vWorkstreamsTable.Columns.Contains("IsHistoric") Then
        vRows = vWorkstreamsTable.Select("IsHistoric='N'")
      End If
      For Each vRow As DataRow In vRows
        vWorkstreamGroupList.Add(vRow("WorkstreamGroup").ToString(), vRow("WorkstreamGroupDesc").ToString())
      Next
    End If

    'Create the menu
    If vWorkstreamGroupList.Count > 0 Then
      Dim vWorkstreamMenu As New ToolStripMenuItem(ControlText.MnuSWorkstreams)
      pParentMenu.DropDownItems.Add(vWorkstreamMenu)

      For Each vItem As DictionaryEntry In vWorkstreamGroupList
        Dim vGroup As ToolStripItem = vWorkstreamMenu.DropDownItems.Add(vItem.Value.ToString)
        vGroup.Tag = vItem.Key
        AddHandler vGroup.Click, AddressOf LaunchWorkstreamGroupFromMenu
        If FormView = FormViews.Modern Then
          Dim vMenuCommand As New MenuToolbarCommand(vItem.Key.ToString, vItem.Value.ToString, CommandIndexes.mnuWorkstreamGroup)
          vMenuCommand.ExplorerMenuAttribute = New ExplorerMenuAttribute(ExplorerMenuSection.System, ExplorerMenuCategory.Workstreams)
          vMenuCommand.OnClick = AddressOf LaunchWorkstreamGroupFromCommand
          vMenuCommand.SerialisationFormat = String.Format("WST.{0}", vItem.Key.ToString())
          mvMenuItems.Add(vItem.Value.ToString, vMenuCommand)
        End If
      Next
    End If

  End Sub

  Private Sub LaunchWorkstreamGroupFromMenu(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim vItem As ToolStripItem = DirectCast(sender, ToolStripItem)
    Dim vWorkstreamGroup As String = vItem.Tag.ToString

    LaunchWorkstreamGroup(vWorkstreamGroup)
  End Sub

  Private Sub LaunchWorkstreamGroupFromCommand(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vItem As MenuToolbarCommand = DirectCast(sender, MenuToolbarCommand)
    Dim vWorkstreamGroup As String = vItem.CommandName

    LaunchWorkstreamGroup(vWorkstreamGroup)
  End Sub

  Private Sub LaunchWorkstreamGroup(pWorkstreamGroup As String)
    FormHelper.ShowWorkstreamIndex(pWorkstreamGroup)
  End Sub
End Class
