Imports CDBNETCL.BrowserMenuHelper
Imports Advanced.Client
Imports CDBNETXAML
Imports CDBNETCL

Public Class BrowserMenu
  Inherits ContextMenuStrip

  Private mvParent As MaintenanceParentForm
  Private mvEntityType As HistoryEntityTypes
  Private mvNumber As Integer
  Private mvDescription As String
  Private mvGroupCode As String
  Private mvNumberList As New ArrayListEx
  Private mvContactInfo As ContactInfo
  Private mvSelectedContacts As List(Of ContactInfo)
  Private mvRemoveSupported As Boolean
  Private mvFavourite As Boolean

  Public Event Remove(ByVal pEntityType As HistoryEntityTypes, ByVal pNumber As Integer)

  Private mvMainMenuItems As New CollectionList(Of MenuToolbarCommand)
  Private mvActivityList As CollectionList(Of String)

  Private Property ContactInfo As ContactInfo
    Get
      If mvContactInfo Is Nothing Then Me.ContactInfo = CreateContactInfo()
      Return mvContactInfo
    End Get
    Set(value As ContactInfo)
      mvContactInfo = value
    End Set
  End Property

  Private Function CreateContactInfo() As ContactInfo
    Dim vResult As ContactInfo = Nothing
    If Me.EntityType = HistoryEntityTypes.hetContacts AndAlso Me.ItemNumber > 0 Then
      vResult = New ContactInfo(Me.ItemNumber)
    End If
    Return vResult
  End Function

  Private Enum RelationshipsMenuItems
    rmiRelationshipsNetworkTo
    rmiRelationshipsNetworkFrom
    rmiRelationshipsTo
    rmiRelationshipsFrom
    rmiRelationshipsExtra
  End Enum

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent
    With mvMainMenuItems
      .Add(BrowserMenuItems.bmiNew.ToString, New WeightedCommand(Of BrowserMenuItems)("New", ControlText.MnuBrowserNew, BrowserMenuItems.bmiNew, "SCBMNE"))
      .Add(BrowserMenuItems.bmiEdit.ToString, New WeightedCommand(Of BrowserMenuItems)("Edit", ControlText.MnuBrowserEdit, BrowserMenuItems.bmiEdit, "SCBMEE"))
      .Add(BrowserMenuItems.bmiEditDetails.ToString, New WeightedCommand(Of BrowserMenuItems)("EditDetails", ControlText.MnuBrowserEditDetails, BrowserMenuItems.bmiEditDetails, "SCBMED"))
      .Add(BrowserMenuItems.bmiDetails.ToString, New WeightedCommand(Of BrowserMenuItems)("Details", ControlText.MnuBrowserDetails, BrowserMenuItems.bmiDetails, "SCBMDE"))
      .Add(BrowserMenuItems.bmiDetailsNewWindow.ToString, New WeightedCommand(Of BrowserMenuItems)("DetailsNewWindow", ControlText.MnuBrowserDetailsNewWindow, BrowserMenuItems.bmiDetailsNewWindow, "SCBMDE"))
      .Add(BrowserMenuItems.bmiSticky.ToString, New WeightedCommand(Of BrowserMenuItems)("Sticky", ControlText.MnuBrowserSticky, BrowserMenuItems.bmiSticky, "SCBMSN"))
      .Add(BrowserMenuItems.bmiJournal.ToString, New WeightedCommand(Of BrowserMenuItems)("Journal", ControlText.MnuBrowserJournal, BrowserMenuItems.bmiJournal, "SCBMJO"))
      .Add(BrowserMenuItems.bmiSendEMail.ToString, New WeightedCommand(Of BrowserMenuItems)("SendEMail", ControlText.MnuBrowserSendEMail, BrowserMenuItems.bmiSendEMail, "SCBMSE"))
      .Add(BrowserMenuItems.bmiDialNumber.ToString, New WeightedCommand(Of BrowserMenuItems)("DialNumber", ControlText.MnuBrowserDialNumber, BrowserMenuItems.bmiDialNumber, "SCBMDN"))
      .Add(BrowserMenuItems.bmiSuppressions.ToString, New WeightedCommand(Of BrowserMenuItems)("Suppressions", ControlText.MnuBrowserSuppressions, BrowserMenuItems.bmiSuppressions, "SCBMSU"))
      .Add(BrowserMenuItems.bmiCommunications.ToString, New WeightedCommand(Of BrowserMenuItems)("Communications", ControlText.MnuBrowserCommunications, BrowserMenuItems.bmiCommunications, "SCBMCM"))
      .Add(BrowserMenuItems.bmiActions.ToString, New WeightedCommand(Of BrowserMenuItems)("Actions", ControlText.MnuBrowserActions, BrowserMenuItems.bmiActions, "SCBMAC"))
      .Add(BrowserMenuItems.bmiActivities.ToString, New WeightedCommand(Of BrowserMenuItems)("Activities", ControlText.MnuBrowserActivities, BrowserMenuItems.bmiActivities, "SCBMAT"))
      .Add(BrowserMenuItems.bmiRelationships.ToString, New WeightedCommand(Of BrowserMenuItems)("Relationships", ControlText.MnuBrowserRelationships, BrowserMenuItems.bmiRelationships, "SCBMRE"))
      .Add(BrowserMenuItems.bmiReport.ToString, New WeightedCommand(Of BrowserMenuItems)("Report", ControlText.MnuSelectionSetReport, BrowserMenuItems.bmiReport, "SCBMRP"))
      .Add(BrowserMenuItems.bmiMailing.ToString, New WeightedCommand(Of BrowserMenuItems)("Mailing", ControlText.MnuSelectionSetMailing, BrowserMenuItems.bmiMailing, "SCBMMA"))
      .Add(BrowserMenuItems.bmiReports.ToString, New WeightedCommand(Of BrowserMenuItems)("Reports", ControlText.MnuBrowserReports, BrowserMenuItems.bmiReports, "SCBMRS"))
      .Add(BrowserMenuItems.bmiSetStatus.ToString, New WeightedCommand(Of BrowserMenuItems)("SetStatus", ControlText.MnuBrowserSetStatus, BrowserMenuItems.bmiSetStatus, "SCBMSS"))
      .Add(BrowserMenuItems.bmiActionSchedule.ToString, New WeightedCommand(Of BrowserMenuItems)("ActionSchedule", ControlText.MnuBrowserActionSchedule, BrowserMenuItems.bmiActionSchedule, "SCBMAS"))
      .Add(BrowserMenuItems.bmiDeleteAllContacts.ToString, New WeightedCommand(Of BrowserMenuItems)("DeleteAllContacts", ControlText.MnuSelectionSetDeleteAllContacts, BrowserMenuItems.bmiDeleteAllContacts, "SCSSDA"))
      .Add(BrowserMenuItems.bmiConvertToOrganisation.ToString, New WeightedCommand(Of BrowserMenuItems)("ConvertToOrganisation", ControlText.MnuBrowserConvertoToOrganisation, BrowserMenuItems.bmiConvertToOrganisation, "SCBMCO"))
      .Add(BrowserMenuItems.bmiCloneOrganisation.ToString, New WeightedCommand(Of BrowserMenuItems)("CloneOrganisation", ControlText.MnuBrowserCloneOrganisation, BrowserMenuItems.bmiCloneOrganisation, "SCBMOC"))
      .Add(BrowserMenuItems.bmiRemove.ToString, New WeightedCommand(Of BrowserMenuItems)("Remove", ControlText.MnuBrowserRemoveItem, BrowserMenuItems.bmiRemove, "SCBMRI"))
      .Add(BrowserMenuItems.bmiAddToFavourites.ToString, New WeightedCommand(Of BrowserMenuItems)("AddToFavourites", ControlText.MnuBrowserAddToFavourites, BrowserMenuItems.bmiAddToFavourites, "SCBMAF"))
      .Add(BrowserMenuItems.bmiNewDocument.ToString, New WeightedCommand(Of BrowserMenuItems)("NewDocument", ControlText.MnuMNewDocument, BrowserMenuItems.bmiNewDocument, "SCFLND"))
      .Add(BrowserMenuItems.bmiSaveAsSelectionSet.ToString, New WeightedCommand(Of BrowserMenuItems)("SaveAsSelectionSet", ControlText.MnuMSaveAsSelectionSet, BrowserMenuItems.bmiSaveAsSelectionSet, "SCSSSS"))
      .Add(BrowserMenuItems.bmiGoToListManager.ToString, New WeightedCommand(Of BrowserMenuItems)("GoToListManager", ControlText.MnuMGoToListManager, BrowserMenuItems.bmiGoToListManager, "SCSSLM"))
      .Add(BrowserMenuItems.bmiDelete.ToString, New WeightedCommand(Of BrowserMenuItems)("Delete", ControlText.MnuSelectionSetDelete, BrowserMenuItems.bmiDelete, "SCSSDE"))
      .Add(BrowserMenuItems.bmiMerge.ToString, New WeightedCommand(Of BrowserMenuItems)("Merge", ControlText.MnuSelectionSetMerge, BrowserMenuItems.bmiMerge, "SCSSME"))
      .Add(BrowserMenuItems.bmiRename.ToString, New WeightedCommand(Of BrowserMenuItems)("Rename", ControlText.MnuSelectionSetRename, BrowserMenuItems.bmiRename, "SCSSRN"))
      .Add(BrowserMenuItems.bmiCopy.ToString, New WeightedCommand(Of BrowserMenuItems)("Copy", ControlText.MnuSelectionSetCopy, BrowserMenuItems.bmiCopy, "SCSSCP"))
      .Add(BrowserMenuItems.bmiSurveyRegistration.ToString, New WeightedCommand(Of BrowserMenuItems)("SurveyRegistration", ControlText.MnuSurveyRegistration, BrowserMenuItems.bmiSurveyRegistration, "SCSSSR"))
      .Add(BrowserMenuItems.bmiBulkUpdateActivity.ToString, New WeightedCommand(Of BrowserMenuItems)("BulkUpdateActitiy", ControlText.MnuBulkUpdateActivity, BrowserMenuItems.bmiBulkUpdateActivity, "SCSSUA"))
      .Add(BrowserMenuItems.bmiDuplicate.ToString, New WeightedCommand(Of BrowserMenuItems)("Duplicate", ControlText.MnuMeetingDuplicate, BrowserMenuItems.bmiDuplicate, "SCBMEE"))
    End With
    For Each vItem As MenuToolbarCommand In mvMainMenuItems
      Dim vNewItem As ToolStripMenuItem = MenuToolbarCommand.NewMenuItem(vItem.CommandID, vItem.MenuText, AddressOf MainMenuHandler)
      Me.Items.Add(vNewItem)

    Next

    DirectCast(Me.Items(BrowserMenuItems.bmiCommunications), ToolStripMenuItem).DropDownItems.AddRange(New ToolStripMenuItem() {
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiCommsRecent, ControlText.MnuBrowserCommsRecent, AddressOf CommsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiCommsAddressed, ControlText.MnuBrowserCommsAddressed, AddressOf CommsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiCommsSent, ControlText.MnuBrowserCommsSent, AddressOf CommsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiCommsRelated, ControlText.MnuBrowserCommsRelated, AddressOf CommsMenuHandler)
                          })
    DirectCast(Me.Items(BrowserMenuItems.bmiActions), ToolStripMenuItem).DropDownItems.AddRange(New ToolStripMenuItem() {
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiActionsNew, ControlText.MnuBrowserActionsNew, AddressOf ActionsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiActionsNewFromTemplate, ControlText.MnuBrowserActionsNewFromTemplate, AddressOf ActionsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiActionsOutstanding, ControlText.MnuBrowserActionsOutstanding, AddressOf ActionsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiActionsCompleted, ControlText.MnuBrowserActionsCompleted, AddressOf ActionsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiActionsOverdue, ControlText.MnuBrowserActionsOverdue, AddressOf ActionsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(BrowserMenuItems.bmiActionsAll, ControlText.MnuBrowserActionsAll, AddressOf ActionsMenuHandler)
                          })
    DirectCast(Me.Items(BrowserMenuItems.bmiRelationships), ToolStripMenuItem).DropDownItems.AddRange(New ToolStripMenuItem() {
                          MenuToolbarCommand.NewMenuItem(RelationshipsMenuItems.rmiRelationshipsNetworkTo, ControlText.MnuBrowserRelationshipsNetworkTo, AddressOf RelationshipsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(RelationshipsMenuItems.rmiRelationshipsNetworkFrom, ControlText.MnuBrowserRelationshipsNetworkFrom, AddressOf RelationshipsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(RelationshipsMenuItems.rmiRelationshipsTo, ControlText.MnuBrowserRelationshipsTo, AddressOf RelationshipsMenuHandler),
                          MenuToolbarCommand.NewMenuItem(RelationshipsMenuItems.rmiRelationshipsFrom, ControlText.MnuBrowserRelationshipsFrom, AddressOf RelationshipsMenuHandler)
                          })
    AddHandler DirectCast(Me.Items(BrowserMenuItems.bmiRelationships), ToolStripMenuItem).DropDownOpening, AddressOf RelationshipsPopupHandler
    mvEntityType = HistoryEntityTypes.hetNone
    MenuToolbarCommand.SetAccessControl(mvMainMenuItems)
  End Sub

  Public Property EntityType() As HistoryEntityTypes
    Get
      Return mvEntityType
    End Get
    Set(ByVal Value As HistoryEntityTypes)
      mvEntityType = Value
    End Set
  End Property
  Public Property ItemNumber() As Integer
    Get
      Return mvNumber
    End Get
    Set(ByVal Value As Integer)
      mvNumber = Value
    End Set
  End Property
  Public Property ItemList() As ArrayListEx
    Get
      Return mvNumberList
    End Get
    Set(ByVal Value As ArrayListEx)
      mvNumberList = Value
    End Set
  End Property
  Public Property ItemDescription() As String
    Get
      Return mvDescription
    End Get
    Set(ByVal Value As String)
      mvDescription = Value
    End Set
  End Property
  Public Property GroupCode() As String
    Get
      Return mvGroupCode
    End Get
    Set(ByVal Value As String)
      mvGroupCode = Value
    End Set
  End Property
  Public Property RemoveSupported() As Boolean
    Get
      Return mvRemoveSupported
    End Get
    Set(ByVal value As Boolean)
      mvRemoveSupported = value
    End Set
  End Property
  Private Sub MainMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MenuHandler(DirectCast(sender, ToolStripMenuItem), CType(DirectCast(sender, ToolStripMenuItem).Tag, BrowserMenuItems))
  End Sub
  Private Sub CommsMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MenuHandler(DirectCast(sender, ToolStripMenuItem), CType(DirectCast(sender, ToolStripMenuItem).Tag, BrowserMenuItems))
  End Sub
  Private Sub ActionsMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MenuHandler(DirectCast(sender, ToolStripMenuItem), CType(DirectCast(sender, ToolStripMenuItem).Tag, BrowserMenuItems))
  End Sub
  Private Sub ActivityMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Dim vOwner As Form = mvParent
      If vOwner Is Nothing Then vOwner = CurrentMainForm
      BrowserMenuHelper.ProcessActivityMenu(vOwner, Me.ContactInfo, vMenuItem.Tag.ToString, vMenuItem.Text)
      'ShowDataSheet(vOwner, frmDataSheet.DataSheetTypes.dstActivities, mvContactInfo, "B", "", vMenuItem.Tag.ToString, vMenuItem.Text)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub RelationshipsMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try

      Dim vInfo As RelationshipMenuData = TryCast(DirectCast(sender, ToolStripMenuItem).Tag, RelationshipMenuData)
      If vInfo Is Nothing Then
        Dim vMenuIndex As RelationshipsMenuItems = CType(DirectCast(sender, ToolStripMenuItem).Tag, RelationshipsMenuItems)
        Select Case vMenuIndex
          Case RelationshipsMenuItems.rmiRelationshipsNetworkFrom
            Dim vForm As New frmNetWork(Me.ContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom)
            vForm.Show()
          Case RelationshipsMenuItems.rmiRelationshipsNetworkTo
            Dim vForm As New frmNetworkNew(Me.ContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo)
            vForm.Show()
          Case RelationshipsMenuItems.rmiRelationshipsFrom
            FormHelper.MaintainContactData(mvNumber, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom, mvParent)
          Case RelationshipsMenuItems.rmiRelationshipsTo
            FormHelper.MaintainContactData(mvNumber, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, mvParent)
        End Select
      Else
        If vInfo.RelationshipMenuItemType = RelationshipMenuData.RelationshipMenuItemTypes.rmitGroup Then
          Dim vOwner As Form = mvParent
          If vOwner Is Nothing Then vOwner = CurrentMainForm
          ShowDataSheet(vOwner, frmDataSheet.DataSheetTypes.dstRelationships, ContactInfo, "B", "", vInfo.RelationshipGroupCode, DirectCast(sender, ToolStripMenuItem).Text)
        Else
          Dim vList As New ParameterList(True)
          vList.IntegerValue("ContactNumber2") = vInfo.ContactInfo.ContactNumber
          Select Case vInfo.RelationshipMenuItemType
            Case RelationshipMenuData.RelationshipMenuItemTypes.rmitLinkTo
              FormHelper.MaintainContactData(mvNumber, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, mvParent, vList)
            Case RelationshipMenuData.RelationshipMenuItemTypes.rmitLinkFrom
              FormHelper.MaintainContactData(mvNumber, CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom, mvParent, vList)
          End Select
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub ReportMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vReportNumber As Integer = 0
      If TypeOf (sender) Is ToolStripMenuItem Then
        Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        vReportNumber = CInt(vMenuItem.Tag)
      Else
        vReportNumber = IntegerValue(sender)
      End If
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ReportNumber") = vReportNumber
      vList.IntegerValue("RP1") = mvNumber
      vList("RP2") = AppValues.Logname()
      Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub StatusMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vStatusCode As String = String.Empty
      If TypeOf (sender) Is ToolStripMenuItem Then
        Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        vStatusCode = vMenuItem.Tag.ToString
      Else
        vStatusCode = sender.ToString()
      End If
      Dim vList As New ParameterList(True)
      vList("Status") = vStatusCode
      vList("StatusDate") = AppValues.TodaysDate
      Dim vActioner As Integer
      Dim vManager As Integer
      For Each vContactInfo As ContactInfo In mvSelectedContacts
        vList.IntegerValue("ContactNumber") = vContactInfo.ContactNumber
        Dim vReturnList As ParameterList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctContact, vList)
        CheckForActioners(vReturnList, mvParent, vActioner, vManager)
      Next
      If mvParent IsNot Nothing Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctNone)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub MenuHandler(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As BrowserMenuItems)
    Dim vCursor As New BusyCursor
    Try
      Select Case pItem
        Case BrowserMenuItems.bmiNew
          Select Case mvEntityType
            Case HistoryEntityTypes.hetActions
              FormHelper.EditAction(0, mvParent)
            Case HistoryEntityTypes.hetContacts
              Dim vList As New ParameterList
              If DataHelper.ContactAndOrganisationGroups.ContainsKey(mvGroupCode) Then
                Dim vType As ContactInfo.ContactTypes
                If DataHelper.ContactAndOrganisationGroups.Item(mvGroupCode).Type = EntityGroup.EntityGroupTypes.egtOrganisationGroup Then
                  vType = ContactInfo.ContactTypes.ctOrganisation
                  vList("OrganisationGroup") = mvGroupCode
                Else
                  vType = ContactInfo.ContactTypes.ctContact
                  vList("ContactGroup") = mvGroupCode
                End If
                FormHelper.ShowNewContactOrDedup(vType, vList, mvParent)
              End If
            Case HistoryEntityTypes.hetDocuments
              FormHelper.NewDocument(mvParent)
            Case HistoryEntityTypes.hetEvents
              FormHelper.ShowEventIndex(0, mvGroupCode)
            Case HistoryEntityTypes.hetSelectionSets
              FormHelper.AddSelectionSet(mvParent)
            Case HistoryEntityTypes.hetMeetings
              FormHelper.EditMeeting(0)
            Case HistoryEntityTypes.hetWorkstreams
              FormHelper.NewWorkstream(mvGroupCode)
          End Select
        Case BrowserMenuItems.bmiEdit
          Select Case mvEntityType
            Case HistoryEntityTypes.hetActions
              FormHelper.EditAction(mvNumber, mvParent)
            Case HistoryEntityTypes.hetEvents
              FormHelper.ShowEventIndex(mvNumber, , mvParent)
            Case HistoryEntityTypes.hetMeetings
              FormHelper.EditMeeting(mvNumber)
          End Select
        Case BrowserMenuItems.bmiEditDetails
          Select Case mvEntityType
            Case HistoryEntityTypes.hetDocuments
              FormHelper.EditDocument(mvNumber, mvParent)
          End Select
        Case BrowserMenuItems.bmiDetails, BrowserMenuItems.bmiDetailsNewWindow
          Dim vList As New ParameterList
          Select Case mvEntityType
            Case HistoryEntityTypes.hetActions
              'This is the 'Open' menu
              vList("ActionNumbers") = mvNumberList.CSList
              FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftActions, vList, mvParent)
            Case HistoryEntityTypes.hetDocuments
              'This is the 'Open' menu
              vList("DocumentNumbers") = mvNumberList.CSList
              FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDocuments, vList)
            Case HistoryEntityTypes.hetSelectionSets
              FormHelper.ShowSelectionSet(mvNumber, mvDescription)
            Case HistoryEntityTypes.hetContacts
              FormHelper.ShowContactCardIndex(mvNumber, pItem = BrowserMenuItems.bmiDetailsNewWindow)
            Case HistoryEntityTypes.hetWorkstreams
              FormHelper.ShowWorkstreamIndex(mvGroupCode, mvNumber)
          End Select
        Case BrowserMenuItems.bmiSticky
          EditContactData(mvParent, Me.ContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactStickyNotes, False)
        Case BrowserMenuItems.bmiJournal
          FormHelper.ShowCardDisplay(CareServices.XMLContactDataSelectionTypes.xcdtContactJournals, ContactInfo.ContactNumber)
        Case BrowserMenuItems.bmiSendEMail
          If EMailApplication.EmailInterface.CanEMail Then
            If mvEntityType = HistoryEntityTypes.hetSelectionSets OrElse
              mvEntityType = HistoryEntityTypes.hetContacts AndAlso ContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
              Dim vForm As frmSelectItems
              If mvEntityType = HistoryEntityTypes.hetSelectionSets Then
                vForm = New frmSelectItems(mvNumber, False)
              Else
                vForm = New frmSelectItems(Me.ContactInfo)
              End If
              If vForm.ShowDialog = DialogResult.OK Then
                EMailApplication.EmailInterface.SendMail(mvParent, EmailInterface.SendEmailOptions.seoAddressResolveUI Or EmailInterface.SendEmailOptions.seoMultipleRecipients, "", "", vForm.EMailAddresses)
              End If
            Else
              EMailApplication.EmailInterface.SendMail(mvParent, EmailInterface.SendEmailOptions.seoAddressResolveUI Or EmailInterface.SendEmailOptions.seoMultipleRecipients, "", "", ContactInfo.EMailAddresses)
            End If
          End If
        Case BrowserMenuItems.bmiDialNumber
          PhoneApplication.PhoneInterface.DialNumber(Me.ContactInfo)
        Case BrowserMenuItems.bmiSuppressions
          EditContactData(mvParent, Me.ContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactSuppressions, False)
        Case BrowserMenuItems.bmiReport
          Dim vForm As New frmReportDataSelection(mvNumber, False)
          vForm.ShowDialog()
        Case BrowserMenuItems.bmiMailing
          AppHelper.ProcessSelectionSetMailing(MainHelper.MainForm, mvNumber, True)
        Case BrowserMenuItems.bmiActionSchedule
          If mvEntityType = HistoryEntityTypes.hetSelectionSets Then
            Dim vForm As New frmActionSchedule(mvNumber, mvDescription)
            vForm.Show()
          End If
        Case BrowserMenuItems.bmiRemove
          RaiseEvent Remove(mvEntityType, mvNumber)
        Case BrowserMenuItems.bmiDeleteAllContacts
          FormHelper.DoBulkContactDeletion(mvNumber)
        Case BrowserMenuItems.bmiConvertToOrganisation
          Dim vList As New ParameterList(True)
          vList.Add("ContactNumber", mvNumber)
          vList.Add("ConvertToOrganisation", "Y")
          vList = ConvertToOrganisation(vList)
          If vList IsNot Nothing Then
            ShowInformationMessage(InformationMessages.ImConvertedToOrganisation)
            MainHelper.RefreshHistoryData(HistoryEntityTypes.hetContacts, Me.ContactInfo.ContactNumber)
          End If
        Case BrowserMenuItems.bmiCloneOrganisation
          If ShowQuestion(GetInformationMessage(QuestionMessages.QmConfirmCloneOrganisation, DataHelper.ContactAndOrganisationGroups(ContactInfo.ContactGroup).GroupName, ContactInfo.ContactName), MessageBoxButtons.YesNo) = DialogResult.Yes Then
            FormHelper.ShowContactCardIndex(DataHelper.CloneOrganisation(Me.ContactInfo.ContactNumber))
          End If
        Case BrowserMenuItems.bmiSaveAsSelectionSet
          Dim vList As New ParameterList(True)
          vList.Add("FromFavorites", "Y")
          Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, 0, Nothing, vList, Nothing)
          vForm.ShowDialog(Me)
        Case BrowserMenuItems.bmiGoToListManager
          Dim vLM As New ListManager(mvNumber, False)
          vLM.Show()
        Case BrowserMenuItems.bmiDelete
          If Not ConfirmDelete() Then Exit Sub
          Dim vList As ParameterList = New ParameterList(True)
          vList.IntegerValue("SelectionSetNumber") = mvNumber
          DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vList)
          UserHistory.RemoveOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, mvNumber)
          If mvParent IsNot Nothing Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)

        '-----------------------------------------------
        'Communications
        Case BrowserMenuItems.bmiCommsRecent, BrowserMenuItems.bmiCommsAddressed, BrowserMenuItems.bmiCommsSent, BrowserMenuItems.bmiCommsRelated
          Dim vList As New ParameterList
          vList.IntegerValue("ContactNumber") = Me.ContactInfo.ContactNumber
          Select Case pItem
            Case BrowserMenuItems.bmiCommsRecent
              vList("DatedOnOrAfter") = AppValues.TodaysDateAddMonths(-1)
              vList("DatedOnOrBefore") = AppValues.TodaysDate
            Case BrowserMenuItems.bmiCommsAddressed
              vList("DocumentLinkType") = "A"
            Case BrowserMenuItems.bmiCommsSent
              vList("DocumentLinkType") = "S"
            Case BrowserMenuItems.bmiCommsRelated
              vList("DocumentLinkType") = "R"
          End Select
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDocuments, vList)
        '-----------------------------------------------
        'Actions
        Case BrowserMenuItems.bmiActionsNew
          FormHelper.EditAction(0, mvParent, Me.ContactInfo)
        Case BrowserMenuItems.bmiActionsNewFromTemplate
          FormHelper.NewActionFromTemplate(mvParent, Me.ContactInfo.ContactNumber)
        Case BrowserMenuItems.bmiActionsOutstanding, BrowserMenuItems.bmiActionsCompleted, BrowserMenuItems.bmiActionsOverdue, BrowserMenuItems.bmiActionsAll
          Dim vList As New ParameterList
          vList.IntegerValue("ContactNumber") = Me.ContactInfo.ContactNumber
          vList("ActionLinkType") = "R"
          Select Case pItem
            Case BrowserMenuItems.bmiActionsOutstanding
              vList("ActionStatus") = AppValues.ActiveActionStatus
            Case BrowserMenuItems.bmiActionsCompleted
              vList("ActionStatus") = AppValues.CompletedActionStatus
            Case BrowserMenuItems.bmiActionsOverdue
              vList("ActionStatus") = AppValues.OverdueActionStatus
            Case BrowserMenuItems.bmiActionsAll
              'Any
          End Select
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftActions, vList, mvParent)
        '-----------------------------------------------
        'Favourites
        Case BrowserMenuItems.bmiAddToFavourites
          Select Case mvEntityType
            Case HistoryEntityTypes.hetFavourites
              'Do Nothing
            Case HistoryEntityTypes.hetContacts
              UserHistory.AddContactHistoryNode(Me.ContactInfo.ContactNumber, ContactInfo.ContactName, ContactInfo.ContactGroup, True)
            Case HistoryEntityTypes.hetEvents
              UserHistory.AddEventHistoryNode(mvNumber, mvDescription, mvGroupCode, True)
            Case HistoryEntityTypes.hetActions, HistoryEntityTypes.hetDocuments, HistoryEntityTypes.hetSelectionSets
              UserHistory.AddOtherHistoryNode(mvEntityType, mvNumber, mvDescription, True)
            Case HistoryEntityTypes.hetWorkstreams
              UserHistory.AddWorkstreamHistoryNode(mvNumber, mvDescription, mvGroupCode, True)
          End Select
        Case BrowserMenuItems.bmiRename
          Dim vList As New ParameterList(True)
          Dim vParamList As New ParameterList(True)
          vParamList("SelectionSet") = mvNumber.ToString
          Dim vSelectionSetTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSelectionSets, vParamList)
          If Not vSelectionSetTable Is Nothing Then
            vList("SelectionSetDesc") = vSelectionSetTable.Rows(0).Item("SelectionSetDesc").ToString
            vParamList.Remove("SelectionSet")
            'Open Application Parameters form
            vParamList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optRenameSelectionSet, Nothing, vList, "Rename Selection Set")
            If Not vParamList Is Nothing Then
              vParamList("SelectionSetNumber") = mvNumber.ToString
              'Call WebService to Rename SelectionSet
              DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vParamList)
              If mvParent IsNot Nothing Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
              UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, mvNumber, vParamList("SelectionSetDesc").ToString)

            End If
          End If
        Case BrowserMenuItems.bmiCopy
          Dim vList As New ParameterList(True)
          Dim vParamList As New ParameterList(True)
          vParamList("SelectionSet") = mvNumber.ToString
          Dim vSelectionSetTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSelectionSets, vParamList)
          If Not vSelectionSetTable Is Nothing Then
            vList("SelectionSetDesc") = vSelectionSetTable.Rows(0).Item("SelectionSetDesc").ToString
            vParamList.Remove("SelectionSet")
            'Open Application Parameters form
            vParamList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optCopySelectionSet, Nothing, vList, "Copy Selection Set")
            If Not vParamList Is Nothing Then
              vParamList("SelectionSetNumber") = mvNumber.ToString
              'Call WebService to Copy SelectionSet
              Dim vSSData As DataSet = DataHelper.GetTableData(CareNetServices.XMLTableDataSelectionTypes.xtdstSelectionSetData, vParamList)
              If vSSData.Tables("DataRow").Rows.Count > 0 Then
                vList("OldSelectionSetNumber") = mvNumber.ToString
                vList("NumberInMailing") = vSSData.Tables("DataRow").Rows(0).Item("NumberInSet").ToString
                vList("SelectionSetDesc") = vParamList("SelectionSetDesc")
                vList("CopySelectionSet") = "Y"
              End If
              Dim vReturnList As New ParameterList
              vReturnList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vList)
              If mvParent IsNot Nothing Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
              UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, IntegerValue(vReturnList("SelectionSetNumber")), vParamList("SelectionSetDesc").ToString)
            End If
          End If
        Case BrowserMenuItems.bmiMerge
          DataHelper.ClearCachedLookupData()
          mvNumber = FormHelper.MergeSelectionSet(mvParent, mvNumber)
          If mvNumber > 0 Then
            If mvParent IsNot Nothing Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctMergeSelectionSet)
          End If
        Case BrowserMenuItems.bmiSurveyRegistration
          Dim vList As New ParameterList(True)
          vList("SelectionSet") = mvNumber.ToString
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtRegisterSurvey, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case BrowserMenuItems.bmiBulkUpdateActivity
          Dim vList As New ParameterList(True)
          vList("SelectionSet") = mvNumber.ToString
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtBulkUpdateActivity, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
        Case BrowserMenuItems.bmiDuplicate
          FormHelper.DoDuplicateMeeting(mvNumber)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub RelationshipsPopupHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      SetVisibleRelationshipItems()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub BrowserMenu_Closing(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosingEventArgs) Handles Me.Closing
    Try
      If mvEntityType = HistoryEntityTypes.hetContacts Then Me.Items(BrowserMenuItems.bmiCloneOrganisation).Text = ControlText.MnuBrowserCloneOrganisation
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub BrowserMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuVisible As Boolean = SetVisibleItems(sender)
      e.Cancel = (Not vMenuVisible)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Function SetVisibleItems(ByVal sender As Object) As Boolean
    Dim vHasVisibleItems As Boolean = True
    Try
      Dim vShowItems(Me.Items.Count) As Boolean
      Dim vMenuOnGridHeader As Boolean
      If sender IsNot Nothing AndAlso TypeOf (DirectCast(sender, BrowserMenu).SourceControl) Is DisplayGrid Then
        If DirectCast((DirectCast(sender, BrowserMenu).SourceControl), DisplayGrid).MenuClickOnHeader Then
          vMenuOnGridHeader = True
        End If
      End If
      If vMenuOnGridHeader = False Then
        If mvNumber > 0 Then vShowItems(BrowserMenuItems.bmiRemove) = mvRemoveSupported
        If mvFavourite = False Then
          Select Case mvEntityType
            Case HistoryEntityTypes.hetActions, HistoryEntityTypes.hetDocuments,
            HistoryEntityTypes.hetEvents, HistoryEntityTypes.hetSelectionSets
              vShowItems(BrowserMenuItems.bmiAddToFavourites) = mvNumber > 0
          End Select
        End If
        Select Case mvEntityType
          Case HistoryEntityTypes.hetFavourites
            vShowItems(BrowserMenuItems.bmiSaveAsSelectionSet) = True
          Case HistoryEntityTypes.hetActions, HistoryEntityTypes.hetDocuments
            vShowItems(BrowserMenuItems.bmiNew) = mvMainMenuItems(BrowserMenuItems.bmiNewDocument).HideItem = False
            'Allow edit if we have an action or document
            If mvEntityType = HistoryEntityTypes.hetActions Then
              If mvNumber > 0 And (FormHelper.ActionRights(mvNumber) And DataHelper.DocumentAccessRights.darEdit) = DataHelper.DocumentAccessRights.darEdit Then vShowItems(BrowserMenuItems.bmiEdit) = True Else vShowItems(BrowserMenuItems.bmiEdit) = False
            Else
              vShowItems(BrowserMenuItems.bmiEditDetails) = mvNumber > 0
            End If
            'Open menu - do not show when clicking on an individual Action / Document
            If ItemNumber > 0 Then
              vShowItems(BrowserMenuItems.bmiDetails) = False
            Else
              vShowItems(BrowserMenuItems.bmiDetails) = mvNumberList.Count > 0   'Only allow details if we have a list
            End If

          Case HistoryEntityTypes.hetMeetings
            If mvNumber > 0 Then vShowItems(BrowserMenuItems.bmiEdit) = True
            vShowItems(BrowserMenuItems.bmiNew) = True
            If mvNumber > 0 Then vShowItems(BrowserMenuItems.bmiDuplicate) = True
          Case HistoryEntityTypes.hetSelectionSets
            vShowItems(BrowserMenuItems.bmiNew) = True
            Dim vGotSelectionSet As Boolean = mvNumber > 0
            vShowItems(BrowserMenuItems.bmiDetails) = vGotSelectionSet              'Only allow selection of specified selection set
            vShowItems(BrowserMenuItems.bmiSendEMail) = vGotSelectionSet            'Only if there is a specific selection set
            vShowItems(BrowserMenuItems.bmiReport) = vGotSelectionSet               'Only if there is a specific selection set
            vShowItems(BrowserMenuItems.bmiMailing) = vGotSelectionSet               'Only if there is a specific selection set
            vShowItems(BrowserMenuItems.bmiActionSchedule) = vGotSelectionSet       'Only if there is a specific selection set
            vShowItems(BrowserMenuItems.bmiMerge) = vGotSelectionSet
            vShowItems(BrowserMenuItems.bmiRename) = vGotSelectionSet
            vShowItems(BrowserMenuItems.bmiCopy) = vGotSelectionSet
            vShowItems(BrowserMenuItems.bmiSurveyRegistration) = vGotSelectionSet
            'Only if there is a specific selection set and User has access rights 
            vShowItems(BrowserMenuItems.bmiDeleteAllContacts) = vGotSelectionSet
            vShowItems(BrowserMenuItems.bmiGoToListManager) = vGotSelectionSet
            vShowItems(BrowserMenuItems.bmiDelete) = vGotSelectionSet
            vShowItems(BrowserMenuItems.bmiBulkUpdateActivity) = vGotSelectionSet

          Case HistoryEntityTypes.hetEvents
            If DataHelper.EventGroups.ContainsKey(mvGroupCode) Then
              vShowItems(BrowserMenuItems.bmiNew) = DataHelper.EventGroups(mvGroupCode).CanCreate
            Else
              vShowItems(BrowserMenuItems.bmiNew) = True
            End If
            vShowItems(BrowserMenuItems.bmiEdit) = mvNumber > 0                'Only allow selection of specified event
          Case HistoryEntityTypes.hetContacts
            If String.IsNullOrEmpty(mvGroupCode) = False AndAlso DataHelper.ContactAndOrganisationGroups.ContainsKey(mvGroupCode) Then
              vShowItems(BrowserMenuItems.bmiNew) = DataHelper.ContactAndOrganisationGroups(mvGroupCode).CanCreate
            Else
              vShowItems(BrowserMenuItems.bmiNew) = True
            End If
            mvSelectedContacts = New List(Of ContactInfo)
            If mvNumber > 0 Then                                                'If we have a contact
              Try
                Me.ContactInfo = Nothing
                Me.ContactInfo = New ContactInfo(mvNumber)
              Catch vCareEx As CareException
                If vCareEx.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
                  vShowItems(BrowserMenuItems.bmiRemove) = True
                Else
                  DataHelper.HandleException(vCareEx)
                End If
              Catch vEx As Exception
                DataHelper.HandleException(vEx)
              End Try
              If Me.ContactInfo IsNot Nothing Then
                mvSelectedContacts.Add(Me.ContactInfo)
                mvGroupCode = Me.ContactInfo.ContactGroup
                DirectCast(Me.Items(BrowserMenuItems.bmiSetStatus), ToolStripMenuItem).DropDownItems.Clear()
                If mvNumberList Is Nothing OrElse mvNumberList.Count = 0 Then
                  If Me.ContactInfo.OwnershipAccessLevel > ContactInfo.OwnershipAccessLevels.oalBrowse Then   'Check for access rights
                    vShowItems(BrowserMenuItems.bmiDetails) = True
                    vShowItems(BrowserMenuItems.bmiDetailsNewWindow) = True
                    vShowItems(BrowserMenuItems.bmiSticky) = True
                    vShowItems(BrowserMenuItems.bmiJournal) = True
                    vShowItems(BrowserMenuItems.bmiCommunications) = True
                    vShowItems(BrowserMenuItems.bmiActions) = True
                    If Not mvFavourite Then vShowItems(BrowserMenuItems.bmiAddToFavourites) = True 'Check if not favourite
                  End If
                  If Me.ContactInfo.OwnershipAccessLevel >= ContactInfo.OwnershipAccessLevels.oalWrite Then   'Check for access rights
                    vShowItems(BrowserMenuItems.bmiRelationships) = True
                    vShowItems(BrowserMenuItems.bmiSuppressions) = True
                    vShowItems(BrowserMenuItems.bmiConvertToOrganisation) = (Me.ContactInfo.ContactType = ContactInfo.ContactTypes.ctContact)
                    Dim vShowCloneOrg As Boolean = Me.ContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation AndAlso
                                                   Me.ContactInfo.ContactGroup <> EntityGroup.DefaultOrganisationGroupCode
                    vShowItems(BrowserMenuItems.bmiCloneOrganisation) = vShowCloneOrg
                    If vShowCloneOrg Then Me.Items(BrowserMenuItems.bmiCloneOrganisation).Text = String.Format(Me.Items(BrowserMenuItems.bmiCloneOrganisation).Text, DataHelper.ContactAndOrganisationGroups(ContactInfo.ContactGroup).GroupName)
                  End If
                  If Me.ContactInfo.OwnershipAccessLevel > ContactInfo.OwnershipAccessLevels.oalBrowse AndAlso Me.ContactInfo.OwnershipAccessLevel >= AppValues.CommunicationsAccessLevel Then
                    vShowItems(BrowserMenuItems.bmiSendEMail) = True
                    vShowItems(BrowserMenuItems.bmiDialNumber) = True
                  End If
                  DirectCast(Me.Items(BrowserMenuItems.bmiReports), ToolStripMenuItem).DropDownItems.Clear()
                  DirectCast(Me.Items(BrowserMenuItems.bmiActivities), ToolStripMenuItem).DropDownItems.Clear()
                  If DataHelper.ContactAndOrganisationGroups.ContainsKey(mvGroupCode) Then
                    vShowItems(BrowserMenuItems.bmiNew) = DataHelper.ContactAndOrganisationGroups(mvGroupCode).CanCreate
                    If Me.ContactInfo.OwnershipAccessLevel >= ContactInfo.OwnershipAccessLevels.oalWrite Then   'Check for access rights
                      Dim vStatusTransitionsTable As DataTable = DataHelper.ContactAndOrganisationGroups(mvGroupCode).StatusTransitionsTable(mvGroupCode, ContactInfo.Status, True)
                      If vStatusTransitionsTable.Rows.Count > 0 Then
                        vShowItems(BrowserMenuItems.bmiSetStatus) = True
                        For Each vRow As DataRow In vStatusTransitionsTable.Rows
                          DirectCast(Me.Items(BrowserMenuItems.bmiSetStatus), ToolStripMenuItem).DropDownItems.Add(vRow.Item("StatusDesc").ToString, Nothing, AddressOf StatusMenuHandler).Tag = vRow.Item("Status").ToString
                        Next
                      End If
                    End If
                    If Me.ContactInfo.OwnershipAccessLevel > ContactInfo.OwnershipAccessLevels.oalBrowse Then
                      Dim vReports As CollectionList(Of String) = DataHelper.ContactAndOrganisationGroups(mvGroupCode).ReportList
                      If vReports.Count > 0 Then
                        vShowItems(BrowserMenuItems.bmiReports) = True
                        For vIndex As Integer = 0 To vReports.Count - 1
                          DirectCast(Me.Items(BrowserMenuItems.bmiReports), ToolStripMenuItem).DropDownItems.Add(vReports(vReports.ItemKey(vIndex)), Nothing, AddressOf ReportMenuHandler).Tag = vReports.ItemKey(vIndex)
                        Next
                      End If
                    End If
                    If Me.ContactInfo.OwnershipAccessLevel >= ContactInfo.OwnershipAccessLevels.oalWrite Then
                      mvActivityList = DataHelper.ContactAndOrganisationGroups(mvGroupCode).ActivityGroupList
                      If mvActivityList.Count > 0 Then
                        vShowItems(BrowserMenuItems.bmiActivities) = True
                        For vIndex As Integer = 0 To mvActivityList.Count - 1
                          DirectCast(Me.Items(BrowserMenuItems.bmiActivities), ToolStripMenuItem).DropDownItems.Add(mvActivityList(vIndex), Nothing, AddressOf ActivityMenuHandler).Tag = mvActivityList.ItemKey(vIndex)
                        Next
                      End If
                    End If
                  End If
                Else
                  Dim vTransitions As Boolean = True
                  Dim vContactInfo As ContactInfo
                  For Each vNumber As Integer In mvNumberList
                    vContactInfo = New ContactInfo(vNumber)
                    mvSelectedContacts.Add(vContactInfo)
                    If vContactInfo.OwnershipAccessLevel < ContactInfo.OwnershipAccessLevels.oalWrite Then vTransitions = False
                    If vContactInfo.Status <> Me.ContactInfo.Status Then vTransitions = False
                  Next
                  If vTransitions Then SetStatusTransitionsMenu(vShowItems)
                End If
              End If
            End If
          Case HistoryEntityTypes.hetWorkstreams
            vShowItems(BrowserMenuItems.bmiNew) = True
            vShowItems(BrowserMenuItems.bmiDetails) = True
            If Not mvFavourite AndAlso mvNumber > 0 Then vShowItems(BrowserMenuItems.bmiAddToFavourites) = True 'Check if not favourite

        End Select
      End If
      Dim vVisibleCount As Integer
      For vIndex As Integer = 0 To Me.Items.Count - 1
        If vIndex < mvMainMenuItems.Count Then
          If mvMainMenuItems(vIndex).HideItem = False Then
            Me.Items(vIndex).Visible = vShowItems(vIndex)
            mvMainMenuItems(vIndex).IsVisible = vShowItems(vIndex)
            If vShowItems(vIndex) Then vVisibleCount += 1
          Else
            Me.Items(vIndex).Visible = False
            mvMainMenuItems(vIndex).IsVisible = False
          End If
        Else    'Item has been added to the end of the menu - could be the grid menu items
          Me.Items(vIndex).Visible = True
          vVisibleCount += 1
        End If
      Next
      If mvEntityType = HistoryEntityTypes.hetContacts AndAlso mvNumber > 0 AndAlso (vVisibleCount > 0 AndAlso vVisibleCount <= 2) Then
        'If we are showing 1 or 2 menus
        If vShowItems(BrowserMenuItems.bmiNew) = True Then
          If vVisibleCount = 1 OrElse vShowItems(BrowserMenuItems.bmiRemove) = True Then
            'The only menus are New &/or Remove so do not show the New menu
            Me.Items(BrowserMenuItems.bmiNew).Visible = False
            mvMainMenuItems(BrowserMenuItems.bmiNew).IsVisible = False
            vVisibleCount -= 1
          End If
        End If
      End If
      If vVisibleCount = 0 Then vHasVisibleItems = False
      Return vHasVisibleItems
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Function

  Private Sub SetVisibleRelationshipItems()
    Dim pRelationshipsMenu As ToolStripMenuItem = TryCast(Me.Items(BrowserMenuItems.bmiRelationships), ToolStripMenuItem)
    If pRelationshipsMenu IsNot Nothing AndAlso mvMainMenuItems(BrowserMenuItems.bmiRelationships).IsVisible Then
      While pRelationshipsMenu.DropDownItems.Count > BrowserMenu.RelationshipsMenuItems.rmiRelationshipsExtra
        pRelationshipsMenu.DropDownItems.RemoveAt(BrowserMenu.RelationshipsMenuItems.rmiRelationshipsExtra)
      End While
      Dim vMenuItem As ToolStripItem
      Dim vList As New ParameterList(True)
      vList(Me.ContactInfo.ContactGroupParameterName) = mvGroupCode
      vList("UsageCode") = "B"
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRelationshipGroups, vList)
      If Not vTable Is Nothing Then
        For Each vRow As DataRow In vTable.Rows
          vMenuItem = pRelationshipsMenu.DropDownItems.Add(vRow.Item("RelationshipGroupDesc").ToString, Nothing, AddressOf RelationshipsMenuHandler)
          vMenuItem.Tag = New RelationshipMenuData(vRow.Item("RelationshipGroup").ToString)
        Next
      End If
      Dim vCardSet As frmCardSet
      For Each vForm As Form In MainHelper.Forms
        If TypeName(vForm) = GetType(frmCardSet).Name Then
          vCardSet = DirectCast(vForm, frmCardSet)
          If vCardSet.ContactInfo.ContactNumber <> Me.ContactInfo.ContactNumber Then
            vMenuItem = pRelationshipsMenu.DropDownItems.Add(String.Format(ControlText.MnuBrowserRelationshipToContact, vCardSet.ContactInfo.ContactName), Nothing, AddressOf RelationshipsMenuHandler)
            vMenuItem.Tag = New RelationshipMenuData(RelationshipMenuData.RelationshipMenuItemTypes.rmitLinkTo, vCardSet.ContactInfo)
            vMenuItem = pRelationshipsMenu.DropDownItems.Add(String.Format(ControlText.MnuBrowserRelationshipFromContact, vCardSet.ContactInfo.ContactName), Nothing, AddressOf RelationshipsMenuHandler)
            vMenuItem.Tag = New RelationshipMenuData(RelationshipMenuData.RelationshipMenuItemTypes.rmitLinkFrom, vCardSet.ContactInfo)
          End If
        End If
      Next
    End If
  End Sub

  Private Sub SetStatusTransitionsMenu(ByVal pShowItems() As Boolean)
    Dim vStatusTransitionsTable As DataTable = DataHelper.ContactAndOrganisationGroups(mvGroupCode).StatusTransitionsTable(mvGroupCode, Me.ContactInfo.Status, True)
    If vStatusTransitionsTable.Rows.Count > 0 Then
      pShowItems(BrowserMenuItems.bmiSetStatus) = True
      For Each vRow As DataRow In vStatusTransitionsTable.Rows
        DirectCast(Me.Items(BrowserMenuItems.bmiSetStatus), ToolStripMenuItem).DropDownItems.Add(vRow.Item("StatusDesc").ToString, Nothing, AddressOf StatusMenuHandler).Tag = vRow.Item("Status").ToString
      Next
    End If
  End Sub

  Private Sub MenuSelected(ByVal sender As Object, ByVal e As EventArgs)
    Try
      Dim vBrowserItem As BrowserMenuItems
      If TypeOf (sender) Is MenuToolbarCommand Then
        Dim vMenu As MenuToolbarCommand = DirectCast(sender, MenuToolbarCommand)
        vBrowserItem = CType(vMenu.CommandID, BrowserMenuItems)
      Else
        vBrowserItem = CType(sender, BrowserMenuItems)
      End If
      Me.MenuHandler(Nothing, vBrowserItem)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub SetStatusMenuSelected(ByVal sender As Object, ByVal e As String)
    If e.Length > 0 Then
      Dim vStatusCode As String = e
      Me.StatusMenuHandler(vStatusCode, Nothing)
    End If
  End Sub

  Private Sub ReportsMenuSelected(ByVal sender As Object, ByVal e As Integer)
    If e > 0 Then
      Dim vReportNumber As Integer = e
      Me.ReportMenuHandler(vReportNumber, Nothing)
    End If
  End Sub

  Private Class RelationshipMenuData
    Public Enum RelationshipMenuItemTypes
      rmitLinkTo
      rmitLinkFrom
      rmitGroup
    End Enum

    Public RelationshipMenuItemType As RelationshipMenuItemTypes
    Public ContactInfo As ContactInfo
    Public RelationshipGroupCode As String

    Public Sub New(ByVal pType As RelationshipMenuItemTypes, ByVal pContactInfo As ContactInfo)
      RelationshipMenuItemType = pType
      ContactInfo = pContactInfo
    End Sub
    Public Sub New(ByVal pGroupCode As String)
      RelationshipMenuItemType = RelationshipMenuItemTypes.rmitGroup
      RelationshipGroupCode = pGroupCode
    End Sub
  End Class
  Private Function ConvertToOrganisation(ByVal pParams As ParameterList) As ParameterList
    Try
      Return DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctContact, pParams)
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enCannotConvertToOrganisation Then
        ShowErrorMessage(vEx.Message)                 'Cannot Convert to Organisation
      ElseIf vEx.ErrorNumber = CareException.ErrorNumbers.enConfirmConvert Then
        If ShowQuestion(vEx.Message, MessageBoxButtons.YesNo) = DialogResult.Yes Then
          pParams.Add("ConfirmConvert", "Y")
          Return ConvertToOrganisation(pParams)
        End If
      Else
        DataHelper.HandleException(vEx)
      End If
    End Try
    Return Nothing
  End Function
  Public Property Favourite() As Boolean
    Get
      Return mvFavourite
    End Get
    Set(ByVal Value As Boolean)
      mvFavourite = Value
    End Set
  End Property

  Private Sub SetMenuContext(ByVal pUserHistoryItem As UserHistoryItem)
    With pUserHistoryItem
      ItemNumber = .Number
      ItemDescription = .DescriptionDisplay
      GroupCode = .GroupCode
      Favourite = .Favourite
      EntityType = .HistoryEntityType
      Me.ContactInfo = .ContactInfo
      Dim vList As New ArrayListEx
      If (pUserHistoryItem.HistoryEntityType = HistoryEntityTypes.hetActions OrElse pUserHistoryItem.HistoryEntityType = HistoryEntityTypes.hetDocuments) Then
        If pUserHistoryItem.Number > 0 Then
          vList.Add(pUserHistoryItem.Number)
        End If
      End If
      ItemList = vList
    End With
  End Sub

  Function InitialiseContextCommands(pUserHistoryItem As UserHistoryItem) As IContextCommandViewModel
    Dim vRtn As New ContextCommandViewModel()
    SetMenuContext(pUserHistoryItem)
    SetVisibleItems(Nothing)
    SetVisibleRelationshipItems()

    Dim vVisibleItems As List(Of WeightedCommand(Of BrowserMenuItems)) = mvMainMenuItems.
      Cast(Of WeightedCommand(Of BrowserMenuItems)).
    Where(Function(vItem) vItem.IsVisible = True).
    ToList()


    Dim vPrimaryMenu As WeightedCommand(Of BrowserMenuItems) = vVisibleItems.FirstOrDefault(Function(vCommand) vCommand.CommandWeighting = CommandWeighting.Primary)
    If vPrimaryMenu IsNot Nothing Then
      vRtn.PrimaryCommand = New UICommand(pUserHistoryItem.Description, Sub() ExecuteCommand(pUserHistoryItem, vPrimaryMenu))
      vRtn.PrimaryCommand.MajorIconKey = pUserHistoryItem.HistoryEntityType.ToString()
      If pUserHistoryItem.HistoryEntityType = HistoryEntityTypes.hetContacts Then
        If Me.ContactInfo IsNot Nothing Then
          vRtn.PrimaryCommand.MajorIconKey = Me.ContactInfo.ContactType.ToString()
        End If
      End If
      vRtn.PrimaryCommand.AdditionalLabel = String.Format("({0})", pUserHistoryItem.Number.ToString)
      vRtn.PrimaryCommand.Tooltip = vPrimaryMenu.CommandName
    End If

    Dim vSecondaryMenus As List(Of WeightedCommand(Of BrowserMenuItems)) = vVisibleItems.Where(Function(vCommand) vCommand.CommandWeighting = CommandWeighting.Secondary).ToList()
    If vSecondaryMenus IsNot Nothing AndAlso vSecondaryMenus.Count > 0 Then
      For Each vMenu As WeightedCommand(Of BrowserMenuItems) In vSecondaryMenus
        Dim vCommand As New UICommand(vMenu.DisplayText, Sub() ExecuteCommand(pUserHistoryItem, vMenu))
        vCommand.MajorIconKey = vMenu.TypedCommandID.ToString()
        vCommand.Tooltip = vMenu.DisplayText
        vRtn.SecondaryCommands.Add(vCommand)
      Next
    End If

    Dim vTertiaryMenus As List(Of WeightedCommand(Of BrowserMenuItems)) = vVisibleItems.Where(Function(vCommand) vCommand.CommandWeighting = CommandWeighting.Tertiary).ToList()
    If vTertiaryMenus IsNot Nothing AndAlso vTertiaryMenus.Count > 0 Then
      For Each vMenu As WeightedCommand(Of BrowserMenuItems) In vTertiaryMenus
        Dim vCommand As New UICommand(vMenu.DisplayText, Sub() ExecuteCommand(pUserHistoryItem, vMenu))
        vCommand.MajorIconKey = vMenu.TypedCommandID.ToString()
        vCommand.Tooltip = vMenu.DisplayText
        vRtn.TertiaryCommands.Add(vCommand)
      Next
    End If

    Dim vAdditionalMenus As List(Of WeightedCommand(Of BrowserMenuItems)) = vVisibleItems.Where(Function(vCommand) vCommand.CommandWeighting = CommandWeighting.Additional).ToList()
    If vAdditionalMenus IsNot Nothing AndAlso vAdditionalMenus.Count > 0 Then
      For Each vMenu As WeightedCommand(Of BrowserMenuItems) In vAdditionalMenus
        Dim vSubMenu As ToolStripMenuItem = TryCast(Me.Items(vMenu.CommandID), ToolStripMenuItem)
        If vSubMenu IsNot Nothing AndAlso vSubMenu.DropDownItems.Count > 0 Then
          Dim vMultiCommand As New Node(Of String, IUICommand)
          vMultiCommand.Header = vMenu.DisplayText
          For Each vItem As ToolStripMenuItem In vSubMenu.DropDownItems
            Dim vCommand As New UICommand(vItem.Text.Replace("&"c, ""), Sub() vItem.PerformClick())
            vCommand.MajorIconKey = vMenu.TypedCommandID.ToString()
            vMultiCommand.Items.Add(vCommand)
          Next
          vRtn.AdditionalCommands.Add(vMultiCommand)
        End If
      Next
    End If
    Return vRtn

  End Function

  Private Sub ExecuteCommand(pHistoryItem As UserHistoryItem, pCommand As WeightedCommand(Of BrowserMenuItems))
    SetMenuContext(pHistoryItem)
    If pCommand IsNot Nothing Then
      MenuHandler(Nothing, pCommand.TypedCommandID)
    End If
  End Sub

End Class