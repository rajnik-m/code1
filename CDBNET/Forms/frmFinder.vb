
Public Class frmFinder
  Inherits CDBNETCL.frmFinder

  Private WithEvents mvDocumentMenu As BaseDocumentMenu
  Private mvActionMenu As ActionMenu
  Private mvSelectionSetMenu As SelectionSetMenu
  Private WithEvents mvBrowserMenu As BrowserMenu
  Private WithEvents mvFinancialFinderMenu As FinancialFinderMenu
  Private WithEvents mvMailingDocumentMenu As MailingDocumentMenu
  Private WithEvents mvEventFinderMenu As EventFinderMenu
  Private WithEvents mvCustomiseMenu As CustomiseMenu
  Private WithEvents mvMeetingFinderMenu As MeetingFinderMenu
  Private mvSuppressBatchView As Boolean
  Private mvQBEQueryID As Integer
  Private mvQBEQueryName As String = ""

  Public Sub New(ByVal pType As CareServices.XMLDataFinderTypes, ByVal pList As ParameterList)
    MyBase.New(pType, pList, False)
  End Sub

  Public Sub New(ByVal pType As CareServices.XMLDataFinderTypes, ByVal pList As ParameterList, pAllContactGroups As Boolean)
    MyBase.New(pType, pList, pAllContactGroups)
  End Sub

  Protected Overrides Sub InitialiseControls(ByVal pType As CareServices.XMLDataFinderTypes, ByVal pList As ParameterList, ByVal pAllContactGroups As Boolean)
    MyBase.InitialiseControls(pType, pList, pAllContactGroups)
    Select Case mvFinderType
      Case CareServices.XMLDataFinderTypes.xdftActions, CareServices.XMLDataFinderTypes.xdftDocuments
        If mvFinderType = CareServices.XMLDataFinderTypes.xdftDocuments Then
          mvDocumentMenu = New DocumentMenu(Me)
          cmdNew.Visible = mvDocumentMenu.CanAddNew
        Else
          mvActionMenu = New ActionMenu(Me)
        End If
      Case CareServices.XMLDataFinderTypes.xdftOpenBatches
        If pList IsNot Nothing AndAlso pList.ContainsKey("AllowNew") AndAlso pList("AllowNew") = "N" Then cmdNew.Visible = False
        dgrResults.MultipleSelect = True
      Case CareServices.XMLDataFinderTypes.xdftBatches
        If pList IsNot Nothing Then
          If pList.Contains("SuppressBatchView") Then mvSuppressBatchView = True
          If pList.Contains("TraderApplication") AndAlso pList("TraderApplication").Length > 0 Then mvSuppressBatchView = True
          If pList.ContainsKey("AllowNew") AndAlso pList("AllowNew") = "N" Then cmdNew.Visible = False
        End If
      Case CareServices.XMLDataFinderTypes.xdftContacts, CareServices.XMLDataFinderTypes.xdftOrganisations, _
           CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts, CareServices.XMLDataFinderTypes), _
           CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleOrganisations, CareServices.XMLDataFinderTypes)
        mvBrowserMenu = New BrowserMenu(Me)
        mvBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
        If mvIsQueryByExample Then GetDefaultQuery()

      Case CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents
        mvEventFinderMenu = New EventFinderMenu(Me)
        GetDefaultQuery()
      Case CareServices.XMLDataFinderTypes.xdftSelectionSets
        mvSelectionSetMenu = New SelectionSetMenu(Me)
      Case CareServices.XMLDataFinderTypes.xdftTransactions
        mvFinancialFinderMenu = New FinancialFinderMenu(Me)
        dgrResults.AutoSetRowHeight = True
      Case CareServices.XMLDataFinderTypes.xdftContactMailingDocuments, CareNetServices.XMLDataFinderTypes.xdftMailings
        mvMailingDocumentMenu = New MailingDocumentMenu(Me)
      Case CareServices.XMLDataFinderTypes.xdftTextSearch
        dgrResults.AutoSetRowHeight = True
      Case CareServices.XMLDataFinderTypes.xdftEvents
        mvEventFinderMenu = New EventFinderMenu(Me)
        dgrResults.MultipleSelect = True
      Case CareServices.XMLDataFinderTypes.xdftMeetings
        mvMeetingFinderMenu = New MeetingFinderMenu(Me)
    End Select
    InitCustomiseMenu()
  End Sub

  Private Sub InitCustomiseMenu()
    mvCustomiseMenu = New CustomiseMenu
    If mvQBEEditPanels IsNot Nothing Then
      For Each vEPL As EditPanel In mvQBEEditPanels
        Dim vCustomiseMenu As New CustomiseMenu
        AddHandler vCustomiseMenu.UpdatePanel, AddressOf UpdatePanel
        vCustomiseMenu.SetContext(mvFinderType, mvAllContactGroups, GetGroupCode, vEPL.Name)
        vEPL.ContextMenuStrip = vCustomiseMenu
      Next
    Else
      epl.ContextMenuStrip = mvCustomiseMenu
      mvCustomiseMenu.SetContext(mvFinderType, mvAllContactGroups, GetGroupCode)
    End If
  End Sub

  Protected Overrides Sub SetContextMenu()
    MyBase.SetContextMenu()
    Select Case mvFinderType
      Case CareServices.XMLDataFinderTypes.xdftActions
        If IsMainFinder Then
          dgrResults.ContextMenuStrip = mvActionMenu
          dgr2.ContextMenuStrip = ctxMenuStrip
          dgr3.ContextMenuStrip = ctxMenuStrip2
        End If
      Case CareServices.XMLDataFinderTypes.xdftContacts, CareServices.XMLDataFinderTypes.xdftOrganisations, _
           CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts, CareServices.XMLDataFinderTypes), _
           CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleOrganisations, CareServices.XMLDataFinderTypes), _
           CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents, CareServices.XMLDataFinderTypes)
        If IsMainFinder Then
          Dim vMenu As ContextMenuStrip
          If mvFinderType = CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents, CareServices.XMLDataFinderTypes) Then
            vMenu = mvEventFinderMenu
          Else
            vMenu = mvBrowserMenu
            If mvIsQueryByExample Then mvBrowserMenu.RemoveSupported = True
          End If
          dgrResults.ContextMenuStrip = vMenu
          MainHelper.RegisterForNavigation(dgrResults)
          If mvIsQueryByExample Then
            dgrResults.AddGridMenu(vMenu)
            If vMenu.Items.ContainsKey("Grid") Then
              Dim vItem As ToolStripMenuItem = CType(vMenu.Items("Grid"), ToolStripMenuItem)
              If Not mvFinderType = CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents Then
                If Not vItem.DropDownItems.ContainsKey("SaveAsSelectionSet") Then
                  Dim vSaveAsSelectionSetItem As New ToolStripMenuItem(ControlText.MnuMSaveAsSelectionSet, AppHelper.ImageProvider.NewOtherImages32.Images("SaveAsSelectionSet"), AddressOf SaveAsSelectionSetHandler, "SaveAsSelectionSet")
                  vItem.DropDownItems.Add(vSaveAsSelectionSetItem)
                End If
                If Not vItem.DropDownItems.ContainsKey("GoToListManager") Then
                  Dim vGoToListManager As New ToolStripMenuItem(ControlText.MnuMGoToListManager, AppHelper.ImageProvider.NewOtherImages32.Images("GoToListManager"), AddressOf GoToListManagerHandler, "GoToListManager")
                  vItem.DropDownItems.Add(vGoToListManager)
                End If
                If Not vItem.DropDownItems.ContainsKey("ReportMailmerge") Then
                  Dim vReportMailmerge As New ToolStripMenuItem(ControlText.MnuMReportMailmerge, AppHelper.ImageProvider.NewOtherImages32.Images("ReportMailmerge"), AddressOf ReportMailmergeHandler, "ReportMailmerge")
                  vItem.DropDownItems.Add(vReportMailmerge)
                End If
                If Not vItem.DropDownItems.ContainsKey("SendEmail") Then
                  Dim vSendEmail As New ToolStripMenuItem(ControlText.MnuMSendEmail, AppHelper.ImageProvider.NewOtherImages32.Images("SendEmail"), AddressOf SendEmailHandler, "SendEmail")
                  vItem.DropDownItems.Add(vSendEmail)
                End If
                If Not vItem.DropDownItems.ContainsKey("Mailing") Then
                  Dim vMailing As New ToolStripMenuItem(ControlText.MnuSelectionSetMailing, AppHelper.ImageProvider.ListManagerImages32.Images("MailList"), AddressOf MailingHandler, "Mailing")
                  vItem.DropDownItems.Add(vMailing)
                End If
              End If
            End If
            dgrResults.SetToolBarVisible()
          End If
        End If
        bpl.RepositionButtons()
      Case CareServices.XMLDataFinderTypes.xdftDocuments
        If IsMainFinder Then
          dgrResults.ContextMenuStrip = mvDocumentMenu
          dgr2.ContextMenuStrip = ctxMenuStrip
          dgr3.ContextMenuStrip = ctxMenuStrip2
        End If
      Case CareServices.XMLDataFinderTypes.xdftSelectionSets
        If IsMainFinder Then
          dgrResults.ContextMenuStrip = mvSelectionSetMenu
        End If
      Case CareServices.XMLDataFinderTypes.xdftTransactions
        If IsMainFinder Then
          If epl.GetValue("FindTransactionType_P") = "P" Then
            dgrResults.ContextMenuStrip = mvFinancialFinderMenu
          Else
            dgrResults.ContextMenuStrip = Nothing
          End If
        End If
      Case CareServices.XMLDataFinderTypes.xdftContactMailingDocuments
        If epl.GetValue("None_F") = "F" Then dgrResults.ContextMenuStrip = mvMailingDocumentMenu Else dgrResults.ContextMenuStrip = Nothing
      Case CareServices.XMLDataFinderTypes.xdftEvents
        dgrResults.ContextMenuStrip = mvEventFinderMenu
      Case CareServices.XMLDataFinderTypes.xdftMeetings
        dgrResults.ContextMenuStrip = mvMeetingFinderMenu
      Case CareNetServices.XMLDataFinderTypes.xdftMailings
        dgrResults.ContextMenuStrip = mvMailingDocumentMenu
    End Select
  End Sub

  Protected Overrides Sub ProcessNew(ByVal pUsePhoneBook As Boolean, ByVal pAlwaysNew As Boolean)
    Select Case mvFinderType
      Case CareServices.XMLDataFinderTypes.xdftActions
        FormHelper.EditAction(0, Me)
      Case CareServices.XMLDataFinderTypes.xdftDocuments
        FormHelper.NewDocument(Me)
      Case CareServices.XMLDataFinderTypes.xdftCampaigns
        FormHelper.ShowCampaignIndex("", Me, GetCampaignRestrictions())
      Case CareServices.XMLDataFinderTypes.xdftEvents
        FormHelper.ShowEventIndex(0, mvList("EventGroup"))
      Case CareServices.XMLDataFinderTypes.xdftStandardDocuments
        EditStandardDocument("", Me)
      Case CType(CareNetServices.XMLDataFinderTypes.xdftMeetings, CareServices.XMLDataFinderTypes)
        EditMeetings("", Me)
      Case Else
        Dim vContactInfo As ContactInfo
        Dim vList As ParameterList
        If mvList Is Nothing Then
          vList = New ParameterList
        Else
          vList = mvList
          If Not String.IsNullOrEmpty(mvDefaultSource) Then mvList("Source") = mvDefaultSource
        End If
        If pUsePhoneBook Then
          GetPhoneBookContact(vList)
        Else
          epl.AddValuesToList(vList)
          vList.RemoveWildcard()
        End If
        If Not Me.Owner Is Nothing Then Me.Hide()
        Select Case mvFinderType
          Case CareServices.XMLDataFinderTypes.xdftContacts
            If vList.ContainsKey("CreateAtOrganisationNumber") Then    'Being called from add position at organisation
              vContactInfo = New ContactInfo(ContactInfo.ContactTypes.ctContact, vList("ContactGroup"))
              vContactInfo.CreateAtOrganisationNumber = vList.IntegerValue("CreateAtOrganisationNumber")
              mvSelectedItem = FormHelper.ShowNewContact(vContactInfo, vList, Me.Owner)
              If mvSelectedItem > 0 AndAlso vContactInfo.ContactCreated Then mvNewContactAtOrg = True
            Else
              If AppValues.ConfigurationValue(AppValues.ConfigurationValues.uniserv_mail).Length > 0 Then pAlwaysNew = True
              mvSelectedItem = FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctContact, vList, Me.Owner, pUsePhoneBook Or pAlwaysNew)
            End If
          Case CareServices.XMLDataFinderTypes.xdftOrganisations
            mvSelectedItem = FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctOrganisation, vList, Me.Owner)
          Case CareServices.XMLDataFinderTypes.xdftDuplicateContacts
            mvSelectedItem = FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctContact, vList, Me.Owner, True)
          Case CareServices.XMLDataFinderTypes.xdftDuplicateOrganisations
            mvSelectedItem = FormHelper.ShowNewContactOrDedup(ContactInfo.ContactTypes.ctOrganisation, vList, Me.Owner, True)
          Case CareServices.XMLDataFinderTypes.xdftBatches
            mvSelectedItem = 0   'This will create new Batch from outside of here
            DialogResult = System.Windows.Forms.DialogResult.OK
        End Select
        If mvSelectedItem > 0 Then DialogResult = System.Windows.Forms.DialogResult.OK
    End Select
    Me.Close()
  End Sub

  Protected Overrides Sub SelectRow(ByVal pRow As Integer)
    MyBase.SelectRow(pRow)
    Try
      Select Case mvFinderType
        Case CareServices.XMLDataFinderTypes.xdftActions
          mvActionMenu.ActionNumber = mvActionNumber
          mvActionMenu.MasterActionNumber = mvMasterActionNumber
          mvActionMenu.ActionStatus = dgrResults.GetValue(pRow, "ActionStatus")
          mvActionMenu.RelatedContactNumber = IntegerValue(dgrResults.GetValue(pRow, "ContactNumber"))
          mvActionMenu.SetNotify(dgr3)
        Case CareServices.XMLDataFinderTypes.xdftContacts, CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts, CareServices.XMLDataFinderTypes)
          mvBrowserMenu.ItemNumber = mvContactNumber
          If dgrResults.MultipleRowsSelected Then
            mvBrowserMenu.ItemList = dgrResults.GetSelectedRowIntegers("ContactNumber")
          Else
            mvBrowserMenu.ItemList = Nothing
          End If
        Case CareServices.XMLDataFinderTypes.xdftDocuments
          mvDocumentMenu.DocumentNumber = mvDocumentNumber
          mvDocumentMenu.SetNotifyProcessed(dgr3)
        Case CareServices.XMLDataFinderTypes.xdftDuplicateContacts, CareServices.XMLDataFinderTypes.xdftDuplicateOrganisations
          'Display tab pages for the addresses
          Dim vContactNumber As Integer
          If mvFinderType = CareServices.XMLDataFinderTypes.xdftDuplicateContacts Then
            vContactNumber = CInt(dgrResults.GetValue(pRow, "ContactNumber"))
          Else
            vContactNumber = CInt(dgrResults.GetValue(pRow, "OrganisationNumber"))
          End If
          Dim vAccessLevel As String = dgrResults.GetValue(pRow, "OwnershipAccessLevel")

          If vContactNumber > 0 Then
            dgr2.Clear()
            dgr2.ShowIfEmpty = True
            If vAccessLevel <> "B" Then
              'Read/Write Access or not using OwnershipGroups
              mvDataSet2 = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vContactNumber)
              dgr2.Populate(mvDataSet2)
            End If
            dgr2.ShowIfEmpty = True
            dgr2.DisplayTitle = "Address"
            dgr2.SelectRow("Default", "Yes")    'Always highlight the default address
            SetTabPages(2)
            If (vAccessLevel = "B" And TabResults.Height < 87) Then TabResults.Height = 87
            cmdUseAddress.Enabled = (dgr2.RowCount > 0)
          End If
        Case CareServices.XMLDataFinderTypes.xdftSelectionSets
          mvSelectionSetMenu.SelectionSetNumber = IntegerValue(dgrResults.GetValue(pRow, "SetNumber"))
        Case CareServices.XMLDataFinderTypes.xdftOrganisations, CType(CareNetServices.XMLDataFinderTypes.xdftQueryByExampleOrganisations, CareServices.XMLDataFinderTypes)
          mvBrowserMenu.ItemNumber = mvContactNumber
          If dgrResults.MultipleRowsSelected Then
            mvBrowserMenu.ItemList = dgrResults.GetSelectedRowIntegers("OrganisationNumber")
          Else
            mvBrowserMenu.ItemList = Nothing
          End If
        Case CareServices.XMLDataFinderTypes.xdftTransactions
          'only for processed transactions!
          If epl.GetValue("FindTransactionType_P") = "P" Then
            mvFinancialFinderMenu.SetContext(IntegerValue(dgrResults.GetValue(pRow, "ContactNumber")), IntegerValue(dgrResults.GetValue(pRow, "BatchNumber")), IntegerValue(dgrResults.GetValue(pRow, "TransactionNumber")))
          End If
        Case CareServices.XMLDataFinderTypes.xdftContactMailingDocuments
          If epl.GetValue("None_F") = "F" Then mvMailingDocumentMenu.SetContext(IntegerValue(dgrResults.GetValue(pRow, "FulfillmentNumber")))
        Case CareNetServices.XMLDataFinderTypes.xdftMailings
          mvMailingDocumentMenu.SetContext(mvFinderType, IntegerValue(dgrResults.GetValue(pRow, "EmailJobNumber")), IntegerValue(dgrResults.GetValue(pRow, "MailingNumber")))
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub GetDefaultQuery()
    Dim vList As New ParameterList(True)
    vList("XmlDataType") = "QE"   'QBE Queries
    vList(mvEntityGroup.ParameterName) = mvEntityGroup.Code
    vList("ItemDesc") = "Default"   'Query Name
    vList("Department") = DataHelper.UserInfo.Department
    Dim vQueryRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, vList, False)
    If vQueryRow IsNot Nothing Then
      Dim vXML As String = vQueryRow.Item("ItemXml").ToString
      Dim vParameterList As New ParameterList(vXML)
      SetQBEValuesFromList(vParameterList)
      mvQBEQueryID = IntegerValue(vQueryRow.Item("XmlItemNumber").ToString)
      mvQBEQueryName = vQueryRow.Item("ItemDesc").ToString
      Me.Text = String.Format("{0} ({1}) - {2}", ControlText.FrmQueryByExample, mvEntityGroup.GroupDescription, mvQBEQueryName)
    End If
  End Sub

  Protected Overrides Sub ProcessOpen()
    'Only valid for QBE
    Try
      Dim vDefaults As New ParameterList
      vDefaults("Department") = DataHelper.UserInfo.Department
      vDefaults(mvEntityGroup.ParameterName) = mvEntityGroup.Code
      Dim vList As ParameterList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optQBESelectQuery, Nothing, vDefaults, "Open Query")
      If vList IsNot Nothing Then
        Dim vQueryRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, vList, False)
        If vQueryRow IsNot Nothing Then
          Dim vXML As String = vQueryRow.Item("ItemXml").ToString
          Dim vParameterList As New ParameterList(vXML)
          SetQBEValuesFromList(vParameterList)
          mvQBEQueryID = IntegerValue(vQueryRow.Item("XmlItemNumber").ToString)
          mvQBEQueryName = vQueryRow.Item("ItemDesc").ToString
          Me.Text = String.Format("{0} ({1}) - {2}", ControlText.FrmQueryByExample, mvEntityGroup.GroupDescription, mvQBEQueryName)
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Protected Overrides Sub ProcessSave()
    'Only valid for QBE
    Try
      Dim vList As New ParameterList(True)
      Dim vDefaults As New ParameterList
      vDefaults("QBEQueryName") = mvQBEQueryName
      vList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optQBEQueryName, Nothing, vDefaults, "Save Query As")
      If vList IsNot Nothing Then
        'If the name is the same as the existing one then save using the existing ID
        If vList("QBEQueryName") <> mvQBEQueryName Then
          'If the name is different check if exists
          Dim vCheckList As New ParameterList(True)
          vCheckList("XmlDataType") = "QE"   'QBE Queries
          vCheckList(mvEntityGroup.ParameterName) = mvEntityGroup.Code
          vCheckList("ItemDesc") = vList("QBEQueryName")    'Query Name
          vCheckList("Department") = DataHelper.UserInfo.Department
          Dim vQueries As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtXmlDataItems, vCheckList, False)
          If vQueries IsNot Nothing Then
            If ShowQuestion(QuestionMessages.QmFileExists, MessageBoxButtons.YesNo, vList("QBEQueryName")) = System.Windows.Forms.DialogResult.No Then
              Exit Sub
            End If
          End If
        End If
        Dim vXMLList As New ParameterList(True)
        vXMLList("XmlDataType") = "QE"                  'Query By Example
        vXMLList("ItemDesc") = vList("QBEQueryName")    'Query Name
        vXMLList("Department") = DataHelper.UserInfo.Department
        vXMLList(mvEntityGroup.ParameterName) = mvEntityGroup.Code
        Dim vParamList As New ParameterList
        AddQBEValuesToList(vParamList)
        vXMLList("ItemXml") = vParamList.XMLParameterString
        Dim vReturnList As ParameterList = DataHelper.AddXMLDataItem(vXMLList)
        mvQBEQueryID = IntegerValue(vReturnList("XmlItemNumber").ToString)
        mvQBEQueryName = vList("QBEQueryName")
        Me.Text = String.Format("{0} ({1}) - {2}", ControlText.FrmQueryByExample, mvEntityGroup.GroupDescription, mvQBEQueryName)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Protected Overrides Sub ProcessClear()
    MyBase.ProcessClear()
    mvQBEQueryName = ""
    mvQBEQueryID = 0
    If mvIsQueryByExample Then Me.Text = String.Format("{0} ({1})", ControlText.FrmQueryByExample, mvEntityGroup.GroupDescription)
  End Sub

  Private Sub frmFinder_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    If mvBrowserMenu IsNot Nothing Then mvBrowserMenu.Dispose()
    If mvActionMenu IsNot Nothing Then mvActionMenu.Dispose()
    If mvDocumentMenu IsNot Nothing Then mvDocumentMenu.Dispose()
    If mvFinancialFinderMenu IsNot Nothing Then mvFinancialFinderMenu.Dispose()
    If mvMailingDocumentMenu IsNot Nothing Then mvMailingDocumentMenu.Dispose()
    If mvEventFinderMenu IsNot Nothing Then mvEventFinderMenu.Dispose()
    If mvCustomiseMenu IsNot Nothing Then mvCustomiseMenu.Dispose()
    If mvMeetingFinderMenu IsNot Nothing Then mvMeetingFinderMenu.Dispose()
  End Sub

  Private Sub dgrResults_CampaignSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCampaign As String) Handles dgrResults.CampaignSelected
    If mvFinderType = CareServices.XMLDataFinderTypes.xdftCampaignCollections Then
      Dim vCol As Integer = dgrResults.GetColumn("CollectionNumber")
      If vCol >= 0 Then mvList("CollectionNumber") = dgrResults.GetValue(pRow, vCol)
      mvSelectedItem = IntegerValue(dgrResults.GetValue(pRow, vCol))
      DialogResult = System.Windows.Forms.DialogResult.OK
      Me.Close()
    Else
      FormHelper.ShowCampaignIndex(pCampaign, Me, GetCampaignRestrictions())
    End If
  End Sub
  Private Sub dgrResults_ProductSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pProduct As String) Handles dgrResults.ProductSelected
    If mvFinderType = CareServices.XMLDataFinderTypes.xdftProducts AndAlso Me.Owner Is Nothing Then
      Dim vParams As New CDBNETCL.ParameterList(True)
      vParams.Add("Product", pProduct)
      vParams.Add("MaintenanceTableName", "products")
      Dim vDataSet As DataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstTableMaintenanceData, CareServices.XMLTableDataSelectionTypes), vParams)
      If vDataSet IsNot Nothing Then
        If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Rows.Count = 1 Then
          vParams = New CDBNETCL.ParameterList(True)
          For vIndex As Integer = 0 To vDataSet.Tables("DataRow").Columns.Count - 1
            vParams(ProperName(vDataSet.Tables("DataRow").Columns(vIndex).ColumnName)) = vDataSet.Tables("DataRow").Rows.Item(0).Item(vIndex).ToString
          Next
        End If
      End If
      Dim vForm As New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, "products", vParams, Nothing)
      vForm.Text = ControlText.FrmAmend & "Products"
      If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
        ProcessFind(False, False, False)
      End If
    End If
  End Sub
  Protected Overrides Sub dgrResults_EventSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pEventNumber As Integer)
    Dim vCol As Integer
    Try
      If dgrResults.GetValue(pRow, dgrResults.GetColumn("Template")) = "Y" Then
        If ShowQuestion(QuestionMessages.QmCreateEventFromTemplate, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
          Dim vDuplicateEvent As Boolean
          Dim vEventInfo As New CareEventInfo(pEventNumber, "")
          Dim vDefaults As New ParameterList
          vDefaults("EventDesc") = vEventInfo.EventDescription.Replace("Template", "").Trim
          vDefaults("LongDescription") = vEventInfo.LongDescription.Replace("Template", "").Trim
          vDefaults("StartDate") = AppValues.TodaysDate
          mvSelectedItems = Nothing
          If ShowQuestion(QuestionMessages.QmCreateEventFromTemplateUsePersonnel, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            Dim vPersonnelDefaults As New ParameterList
            vPersonnelDefaults("Duration") = CInt(DateDiff(DateInterval.Day, CDate(vEventInfo.StartDate), CDate(vEventInfo.EndDate)) + 1).ToString
            Dim vFinder As frmFinder = New frmFinder(CareServices.XMLDataFinderTypes.xdftEventPersonnel, vPersonnelDefaults, False)
            If vFinder.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
              If vFinder.SelectedItems.Count > 0 Then
                mvSelectedItems = vFinder.SelectedItems
                vDefaults("StartDate") = vPersonnelDefaults("StartDate")
                vDuplicateEvent = True
              End If
            End If
          Else
            vDuplicateEvent = True
          End If
          If vDuplicateEvent Then
            'Create new Event from Template
            Dim vList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateEvent, vDefaults)
            If vList.Count > 0 Then
              vList.Add("Template", "Y")
              Dim vResult As ParameterList = DataHelper.DuplicateEvent(pEventNumber, vList)
              'BR16770 If event pricing matrix does not exist for current date range, return a message and uncheck allow bookings check box.
              If vResult.Contains("PricingMatrixValid") AndAlso Not BooleanValue(vResult("PricingMatrixValid").ToString) Then
                ShowWarningMessage(InformationMessages.ImEventPricingMatrixInvalid)
              End If
              If vResult.Contains("EventNumber") Then
                If mvSelectedItems IsNot Nothing AndAlso mvSelectedItems.Count > 0 Then
                  vEventInfo = New CareEventInfo(IntegerValue(vResult("EventNumber")), "")
                  For Each vItem As Integer In mvSelectedItems
                    Dim vPersonnelList As New ParameterList(True)
                    Dim vPersonnelContact As ContactInfo = New ContactInfo(vItem)
                    vPersonnelList("ContactNumber") = vItem.ToString
                    vPersonnelList("AddressNumber") = vPersonnelContact.AddressNumber.ToString
                    vPersonnelList("EventNumber") = vEventInfo.EventNumber.ToString
                    vPersonnelList("SessionNumber") = vEventInfo.BaseItemNumber.ToString
                    vPersonnelList("StartDate") = vEventInfo.StartDate.ToString(AppValues.DateFormat)
                    vPersonnelList("StartTime") = vEventInfo.StartTime.ToString(AppValues.TimeFormat)
                    vPersonnelList("EndDate") = vEventInfo.EndDate.ToString(AppValues.DateFormat)
                    vPersonnelList("EndTime") = vEventInfo.EndTime.ToString(AppValues.TimeFormat)
                    DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventPersonnel, vPersonnelList)
                  Next
                End If

                If Me.Owner Is Nothing Then
                  FormHelper.ShowEventIndex(vResult.IntegerValue("EventNumber"))
                Else
                  mvSelectedItem = vResult.IntegerValue("EventNumber")
                  DialogResult = System.Windows.Forms.DialogResult.OK
                  Me.Close()
                End If
              End If
            End If
          End If
          Exit Sub
        End If
      End If
      If IsMainFinder Then
        FormHelper.ShowEventIndex(pEventNumber)
      Else
        Select Case mvFinderType
          Case CareServices.XMLDataFinderTypes.xdftEvents
            vCol = dgrResults.GetColumn("EventNumber")
            If vCol >= 0 Then mvList("EventNumber") = dgrResults.GetValue(pRow, vCol)
            mvSelectedItem = pEventNumber
          Case CareServices.XMLDataFinderTypes.xdftEventBookings
            'Do nothing as mvSelectedItem has already been set in CDBNETCL.frmFinder.dgrResults_RowDoubleClicked
          Case Else
            mvSelectedItem = pEventNumber
        End Select
        DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  Protected Overrides Sub dgrResults_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer)
    Try
      If Me.Owner Is Nothing OrElse mvFindInPhoneBook Then
        If mvFindInPhoneBook Then
          mvSelectedItem = pContactNumber
          ProcessNew(True, False)
        Else
          If mvFinderType = CareServices.XMLDataFinderTypes.xdftTransactions Then
            'Need to navigate to the actual transaction found in the finder
            Dim vType As CareServices.XMLContactDataSelectionTypes = CareServices.XMLContactDataSelectionTypes.xcdtContactUnProcessedTransactions
            If mvList("Posted").Length > 0 Then vType = CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions
            Dim vForm As Form = FormHelper.ShowCardIndex(vType, pContactNumber, False)
            If vForm IsNot Nothing Then
              DirectCast(vForm, frmCardSet).SelectTransaction(mvList.IntegerValue("BatchNumber"), mvList.IntegerValue("TransactionNumber"), 0)
            End If
          Else
            FormHelper.ShowContactCardIndex(pContactNumber)
          End If
        End If
      Else
        Dim vCol As Integer
        Dim vParam As String = ""
        Select Case mvFinderType
          Case CareServices.XMLDataFinderTypes.xdftCovenants
            vParam = "CovenantNumber"
          Case CareServices.XMLDataFinderTypes.xdftCommunications
            vParam = "CommunicationNumber"
          Case CareServices.XMLDataFinderTypes.xdftCreditCardAuthorities
            vParam = "CreditCardAuthorityNumber"
          Case CareNetServices.XMLDataFinderTypes.xdftExamPersonnel
            vParam = "ExamPersonnelId"
          Case CareServices.XMLDataFinderTypes.xdftMembers
            vParam = "MemberNumber"
          Case CareServices.XMLDataFinderTypes.xdftPaymentPlans
            vParam = "PaymentPlanNumber"
          Case CareServices.XMLDataFinderTypes.xdftStandingOrders
            vParam = "BankersOrderNumber"
          Case CareServices.XMLDataFinderTypes.xdftDirectDebits
            vParam = "DirectDebitNumber"
          Case CareNetServices.XMLDataFinderTypes.xdftFundraisingRequestsFinder
            vParam = "RequestNumber"
          Case CareNetServices.XMLDataFinderTypes.xdftGiftAidDeclarations
            vParam = "DeclarationNumber" 'BR19268
          Case CareNetServices.XMLDataFinderTypes.xdftEventPersonnel
            Exit Sub
          Case CareNetServices.XMLDataFinderTypes.xdftCPDCyclePeriodFinder
            vParam = "ContactCPDPeriodNumber"
          Case CareNetServices.XMLDataFinderTypes.xdftCPDPointFinder
            vParam = "ContactCPDPointNumber"
        End Select
        If vParam.Length > 0 Then
          vCol = dgrResults.GetColumn(vParam)
          If vParam = "PaymentPlanNumber" AndAlso vCol < 0 Then vCol = dgrResults.GetColumn("OrderNumber")
          If vCol >= 0 Then mvList(vParam) = dgrResults.GetValue(pRow, vCol)
          If mvFinderType = CareServices.XMLDataFinderTypes.xdftMembers AndAlso mvSelectedItem > 0 Then
            mvList("MembershipNumber") = mvSelectedItem.ToString
          End If
          If mvFinderType = CareServices.XMLDataFinderTypes.xdftFundraisingRequestsFinder AndAlso mvSelectedItem > 0 Then
            mvList("RequestNumber") = mvSelectedItem.ToString
          End If
        End If
        Select Case mvFinderType
          Case CareNetServices.XMLDataFinderTypes.xdftCPDCyclePeriodFinder, CareNetServices.XMLDataFinderTypes.xdftCPDPointFinder,
               CareNetServices.XMLDataFinderTypes.xdftFundraisingRequestsFinder
            'Leave mvSelectedItem as it is
          Case Else
            'Reset mvSelectedItem to be the Contact number
            mvSelectedItem = pContactNumber
        End Select
        DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub dgrResults_RowDoubleClicked(ByVal pSender As Object, ByVal pRow As Integer) Handles dgrResults.RowDoubleClicked
    Dim vCursor As New BusyCursor
    Try
      Select Case mvFinderType
        Case CareServices.XMLDataFinderTypes.xdftTextSearch
          mvSelectedItem = IntegerValue(dgrResults.GetValue(pRow, "ItemNumber"))
          Dim vItemType As String = dgrResults.GetValue(pRow, "ItemType")
          If vItemType.StartsWith("C") Then
            FormHelper.ShowContactCardIndex(mvSelectedItem)
          ElseIf vItemType.StartsWith("E") Then
            FormHelper.ShowEventIndex(mvSelectedItem)
          End If
        Case CareServices.XMLDataFinderTypes.xdftAppealCollections
          mvSelectedItem = IntegerValue(dgrResults.GetValue(pRow, "CollectionNumber"))
        Case CareServices.XMLDataFinderTypes.xdftInternalResources
          Dim vNumber As Integer = CInt(dgrResults.GetValue(pRow, "ResourceNumber"))
          mvSelectedItem = vNumber
          If Me.Owner IsNot Nothing Then
            DialogResult = System.Windows.Forms.DialogResult.OK
          End If
        Case CareServices.XMLDataFinderTypes.xdftLegacies
          If mvList IsNot Nothing Then mvList("LegacyNumber") = dgrResults.GetValue(pRow, "LegacyNumber")
        Case CareServices.XMLDataFinderTypes.xdftSelectionSets
          Dim vNumber As Integer = CInt(dgrResults.GetValue(pRow, "SetNumber"))
          Dim vDesc As String = dgrResults.GetValue(pRow, "Description")
          UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, vNumber, dgrResults.GetValue(pRow, "Description"))
          FormHelper.ShowSelectionSet(vNumber, vDesc)
        Case CareServices.XMLDataFinderTypes.xdftProducts
          If Not mvList Is Nothing Then
            Dim vCol As Integer = dgrResults.GetColumn("Product")
            If vCol >= 0 Then mvList("Product") = dgrResults.GetValue(pRow, vCol)
            mvSelectedItem = 1
          End If
        Case CareServices.XMLDataFinderTypes.xdftBatches, CareServices.XMLDataFinderTypes.xdftOpenBatches
          Dim vCol As Integer = dgrResults.GetColumn("BatchNumber")
          If dgrResults.GetSelectedRowIntegers("BatchNumber").Count > 1 Then
            mvSelectedItems = dgrResults.GetSelectedRowIntegers("BatchNumber")
          Else
            If vCol >= 0 Then mvSelectedItem = IntegerValue(dgrResults.GetValue(pRow, vCol))
          End If
          If mvFinderType = CareServices.XMLDataFinderTypes.xdftOpenBatches Then
            If Not mvSelectedItems Is Nothing Then
              FormHelper.CloseOpenBatch(mvSelectedItems)
              ProcessFind(True, False, False)
            ElseIf mvSelectedItem > 0 Then
              FormHelper.CloseOpenBatch(mvSelectedItem)
              ProcessFind(True, False, False)
            End If
          End If
        Case CareServices.XMLDataFinderTypes.xdftMembers
          Dim vCol As Integer = dgrResults.GetColumn("MembershipNumber")
          If vCol >= 0 Then mvSelectedItem = IntegerValue(dgrResults.GetValue(pRow, vCol))
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ctxMenuStrip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctxMenuEdit.Click, ctxMenuNew.Click
    Dim vForm As frmCardMaintenance
    Dim vCursor As New BusyCursor
    Try
      Dim vContactInfo As New ContactInfo(ContactInfo.ContactTypes.ctContact, "")
      Dim vEdit As Boolean = False
      If sender Is ctxMenuEdit Then vEdit = True
      Dim vDataType As CareServices.XMLContactDataSelectionTypes
      If mvFinderType = CareServices.XMLDataFinderTypes.xdftActions Then
        vDataType = CareServices.XMLContactDataSelectionTypes.xcdtNone
      ElseIf mvFinderType = CareServices.XMLDataFinderTypes.xdftDocuments Then
        vDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments
      End If
      Dim vMaintenanceType As CareServices.XMLMaintenanceControlTypes
      Dim vDataSet As DataSet
      Dim vRow As Integer
      If DirectCast(DirectCast(sender, ToolStripMenuItem).GetCurrentParent, ContextMenuStrip).SourceControl Is dgr2 Then
        vDataSet = mvDataSet2
        vRow = dgr2.CurrentRow
        If mvFinderType = CareServices.XMLDataFinderTypes.xdftActions Then
          vContactInfo.SelectedActionNumber = mvActionNumber
          vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        Else
          vContactInfo.SelectedDocumentNumber = mvDocumentNumber
          vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
        End If
      Else
        vDataSet = mvDataSet3
        vRow = dgr3.CurrentRow
        If mvFinderType = CareServices.XMLDataFinderTypes.xdftActions Then
          vContactInfo.SelectedActionNumber = mvActionNumber
          vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink
        Else
          vContactInfo.SelectedDocumentNumber = mvDocumentNumber
          vMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocumentLink
        End If
      End If
      vForm = New frmCardMaintenance(Me, vContactInfo, vDataType, vDataSet, vEdit, vRow, vMaintenanceType)
      vForm.Show()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub dgr3_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr3.ContactSelected
    If Me.Owner Is Nothing Then FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub dgr3_DocumentSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr3.DocumentSelected
    If Me.Owner Is Nothing Then
      Dim vList As New ParameterList
      vList.IntegerValue("DocumentNumbers") = pContactNumber
      FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDocuments, vList)
    End If
  End Sub

  Private Sub dgr3_EventSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pEventNumber As Integer) Handles dgr3.EventSelected
    If Me.Owner Is Nothing Then
      FormHelper.ShowEventIndex(pEventNumber)
    End If
  End Sub

  Private Sub mvFinancialMenu_MenuSelected(ByVal pItem As FinancialMenu.FinancialMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvFinancialFinderMenu.MenuSelected
    Try
      Select Case pItem
        Case FinancialMenu.FinancialMenuItems.fmiReverse, BaseFinancialMenu.FinancialMenuItems.fmiMove, BaseFinancialMenu.FinancialMenuItems.fmiRefund, BaseFinancialMenu.FinancialMenuItems.fmiAnalysis
          Dim vList As New ParameterList(True, True)
          vList("BatchNumber") = dgrResults.GetValue(dgrResults.CurrentRow, "BatchNumber")
          vList("TransactionNumber") = dgrResults.GetValue(dgrResults.CurrentRow, "TransactionNumber")
          Dim vContactNumber As Integer = IntegerValue(dgrResults.GetValue(dgrResults.CurrentRow, "ContactNumber"))
          'Retrieve ProcessedTransaction data as this includes stock information
          Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions, vContactNumber, vList))
          If vRow IsNot Nothing Then
            vList.Remove("ContactNumber")
            Dim vTransDate As String = vRow.Item("TransactionDate").ToString
            Dim vTransSign As String = vRow.Item("TransactionSign").ToString
            Dim vStock As Boolean = BooleanValue(vRow.Item("ContainsStock").ToString)
            If vStock = False Then vStock = BooleanValue(vRow.Item("ContainsPostage").ToString)
            Select Case pItem
              Case BaseFinancialMenu.FinancialMenuItems.fmiReverse
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atReverse, vList, vTransDate, vTransSign, vStock)
              Case BaseFinancialMenu.FinancialMenuItems.fmiMove
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atMove, vList, vTransDate, vTransSign, vStock)
              Case BaseFinancialMenu.FinancialMenuItems.fmiRefund
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atRefund, vList, vTransDate, vTransSign, vStock)
              Case BaseFinancialMenu.FinancialMenuItems.fmiAnalysis
                FormHelper.RunFinancialAdjustments(CareServices.AdjustmentTypes.atAdjustment, vList, vTransDate, vTransSign, vStock)
            End Select
          End If
      End Select
    Catch vCareException As CareException
      DataHelper.HandleException(vCareException)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Public ReadOnly Property SelectedContactNumber() As Integer
    Get
      Return mvContactNumber
    End Get
  End Property

  Private Sub mvMailingDocumentMenu_MenuSelected(ByVal pItem As BaseFinancialMenu.FinancialMenuItems, ByVal pDataRow As System.Data.DataRow, ByVal pChangeDetails As Boolean, ByVal pFinancialMenu As BaseFinancialMenu) Handles mvMailingDocumentMenu.MenuSelected
    Try
      Select Case pItem
        Case BaseFinancialMenu.FinancialMenuItems.fmiViewMailingDocument
          ViewMailingDocument()
        Case FinancialMenu.FinancialMenuItems.fmiRedoFulfilment
          Dim vNumberOfDocuments As Integer = IntegerValue(dgrResults.GetValue(dgrResults.CurrentRow, "NumberOfDocuments"))
          Dim vDialogResult As DialogResult = System.Windows.Forms.DialogResult.Yes
          If vNumberOfDocuments > 0 Then vDialogResult = ShowQuestion(QuestionMessages.QmConfirmSetUnfulfilled, MessageBoxButtons.YesNo, vNumberOfDocuments.ToString)
          If vDialogResult = System.Windows.Forms.DialogResult.Yes Then
            Dim vResultList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptCancReason, Nothing, Nothing)
            If vResultList IsNot Nothing AndAlso vResultList.Contains("CancellationReason") Then
              DataHelper.SetMailingDocumentUnfulfilled(IntegerValue(dgrResults.GetValue(dgrResults.CurrentRow, "FulfillmentNumber")), vResultList("CancellationReason"))
              dgrResults.DeleteRow(dgrResults.CurrentRow)
            End If
          End If
      End Select
    Catch vCareException As CareException
      DataHelper.HandleException(vCareException)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub mvEventFinderMenu_MenuSelected(ByVal pItem As EventFinderMenu.EventMenuItems) Handles mvEventFinderMenu.MenuSelected
    Try
      If Not dgrResults.MultipleRowsSelected Then
        Dim vEventNumber As Integer = CInt(dgrResults.GetValue(dgrResults.ActiveRow, 0))
        Dim vEventInfo As New CareEventInfo(vEventNumber)
        Dim vDefaults As New ParameterList
        vDefaults("EventDesc") = vEventInfo.EventDescription
        vDefaults("LongDescription") = vEventInfo.LongDescription
        vDefaults("StartDate") = AppValues.TodaysDate
        vDefaults("MultipleRows") = "False"
        Dim vList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateEvent, vDefaults)
        If vList.Count > 0 Then
          vList.Add("Template", CBoolYN(vEventInfo.Template))
          Dim vResult As ParameterList = DataHelper.DuplicateEvent(vEventInfo.EventNumber, vList)
          'BR16770 If event pricing matrix does not exist for current date range, return a message and uncheck allow bookings check box.
          If vResult.Contains("PricingMatrixValid") AndAlso Not BooleanValue(vResult("PricingMatrixValid").ToString) Then
            ShowWarningMessage(InformationMessages.ImEventPricingMatrixInvalid)
          End If
          If vResult.Contains("EventNumber") Then
            vEventInfo = New CareEventInfo(vResult.IntegerValue("EventNumber"), vEventInfo.EventGroup)
            UserHistory.AddEventHistoryNode(vEventInfo.EventNumber, vEventInfo.EventName, vEventInfo.EventGroup)
          End If
        End If
      Else
        Dim vRetainDayOfWeek As Boolean
        Dim vFirstEvent As Boolean = True
        Dim vNumberOfDays As Integer
        Dim vNewStart As DateTime
        Dim vAmendedDays As Boolean
        Dim vCancel As Boolean = False
        Dim vSelectedRows As CDBNETCL.ArrayListEx = dgrResults.GetSelectedRowNumbers
        vSelectedRows.Sort()
        For vRow As Integer = 0 To vSelectedRows.Count - 1
          Dim vStartDate As String = dgrResults.GetValue(CInt(vSelectedRows.Item(vRow)), "StartDate")
          Dim vEventNumber As Integer = CInt(dgrResults.GetValue(CInt(vSelectedRows.Item(vRow)), 0))
          Dim vOldStart As DateTime = CDate(vStartDate)
          Dim day As String = vOldStart.DayOfWeek.ToString
          If vFirstEvent = True Or vRetainDayOfWeek = True AndAlso vAmendedDays = False Then
            vNewStart = vOldStart.AddYears(1)
          Else
            vNewStart = vOldStart.AddDays(vNumberOfDays)
          End If
          If vRetainDayOfWeek = True Or vFirstEvent = True Then
            If vNewStart.DayOfWeek <> vOldStart.DayOfWeek Then
              'get new start day, same as old start day and amend new date
              For vAddDays As Integer = 1 To vAddDays + 4
                Dim vDate As DateTime = vNewStart.AddDays(vAddDays)
                If vDate.DayOfWeek = vOldStart.DayOfWeek Then
                  vNewStart = vDate
                  'get number of days between old date and new date
                  vNumberOfDays = CInt(vNewStart.Subtract(vOldStart).TotalDays)
                  Exit For
                Else
                  vDate = vNewStart.AddDays(-vAddDays)
                  If vDate.DayOfWeek = vOldStart.DayOfWeek Then
                    vNewStart = vDate
                    'get number of days between old date and new date
                    vNumberOfDays = CInt(vNewStart.Subtract(vOldStart).TotalDays)
                    Exit For
                  End If
                End If
              Next
            End If
          End If
          Dim vEventInfo As New CareEventInfo(vEventNumber)
          Dim vDefaults As New ParameterList
          vDefaults("EventDesc") = vEventInfo.EventDescription
          vDefaults("LongDescription") = vEventInfo.LongDescription
          vDefaults("StartDate") = vNewStart.ToString
          vDefaults("MultipleRows") = "True"
          Dim vList As New ParameterList
          Dim vList_1 As ParameterList
          Dim vResult As New ParameterList
          If vFirstEvent = True Then
            vList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateEvent, vDefaults)
          End If
          If vList.Count = 0 AndAlso vFirstEvent = True Then vCancel = True
          If vCancel = False AndAlso vFirstEvent = True Then
            vList.Add("Template", CBoolYN(vEventInfo.Template))
            If vList("StartDate") <> vDefaults("StartDate") Then
              Dim vUserAmendedDate As DateTime = CDate(vList("StartDate").ToString)
              vNewStart = CDate(vUserAmendedDate)
              vNumberOfDays = CInt(vNewStart.Subtract(vOldStart).TotalDays)
              vAmendedDays = True
            End If
            If vList("RetainDayOfWeek") = "Y" Then vRetainDayOfWeek = True
            vList.Remove("RetainDayOfWeek")
            vFirstEvent = False
            vResult = DataHelper.DuplicateEvent(vEventInfo.EventNumber, vList)
            'BR16770 If event pricing matrix does not exist for current date range, return a message and uncheck allow bookings check box.
            If vResult.Contains("PricingMatrixValid") AndAlso Not BooleanValue(vResult("PricingMatrixValid").ToString) Then
              ShowWarningMessage(InformationMessages.ImEventPricingMatrixInvalid)
            End If
          ElseIf vCancel = False AndAlso vFirstEvent = False Then
            vList_1 = New ParameterList(True)
            vList_1("EventDesc") = vDefaults("EventDesc")
            vList_1("StartDate") = vDefaults("StartDate")
            vList_1.Add("Template", CBoolYN(vEventInfo.Template))
            vResult = DataHelper.DuplicateEvent(vEventInfo.EventNumber, vList_1)
            'BR16770 If event pricing matrix does not exist for current date range, return a message and uncheck allow bookings check box.
            If vResult.Contains("PricingMatrixValid") AndAlso Not BooleanValue(vResult("PricingMatrixValid").ToString) Then
              ShowWarningMessage(InformationMessages.ImEventPricingMatrixInvalid)
            End If
          End If
          If vResult.Contains("EventNumber") Then
            vEventInfo = New CareEventInfo(vResult.IntegerValue("EventNumber"), vEventInfo.EventGroup)
          Else : Exit Sub
          End If
        Next
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enAppointmentConflict Then
        ShowInformationMessage(vException.Message)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  Private Sub mvMeetingFinderMenu_MenuSelected(ByVal pItem As MeetingFinderMenu.MeetingMenuItems) Handles mvMeetingFinderMenu.MenuSelected
    Try
      FormHelper.DoDuplicateMeeting(CInt(dgrResults.GetValue(dgrResults.ActiveRow, 0)))
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub UpdatePanel(ByVal pRevert As Boolean) Handles mvCustomiseMenu.UpdatePanel
    Dim vNodeName As String = ""
    If mvIsQueryByExample Then
      vNodeName = mvCurrentQBENodeName
      tabQuery.Visible = False
    Else
      epl.Visible = False
    End If
    MyBase.InitialiseControls(mvFinderType, mvList, mvAllContactGroups)
    InitCustomiseMenu()
    If mvIsQueryByExample Then
      SelectQBENode(vNodeName)
      tabQuery.Visible = True
    Else
      epl.Visible = True
    End If
  End Sub

  Private Function GetGroupCode() As String
    If mvList IsNot Nothing Then
      Select Case mvFinderType
        Case CareServices.XMLDataFinderTypes.xdftContacts, CareNetServices.XMLDataFinderTypes.xdftQueryByExampleContacts
          If mvList.ContainsKey("ContactGroup") Then Return mvList("ContactGroup")
        Case CareServices.XMLDataFinderTypes.xdftOrganisations, CareNetServices.XMLDataFinderTypes.xdftQueryByExampleOrganisations
          If mvList.ContainsKey("OrganisationGroup") Then Return mvList("OrganisationGroup")
        Case CareServices.XMLDataFinderTypes.xdftEvents, CareNetServices.XMLDataFinderTypes.xdftQueryByExampleEvents
          If mvList.ContainsKey("EventGroup") Then Return mvList("EventGroup")
      End Select
    End If
    Return ""
  End Function

  Private Sub SaveAsSelectionSetHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Dim vList As New ParameterList(True)
      vList.Add("FromTemporaryTable", mvQBESelectionSet)
      Dim vSelectionSet As Integer = mvQBESelectionSet
      If SelectionSetExists(vSelectionSet) Then vSelectionSet = 0
      Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vSelectionSet, Nothing, vList, Nothing)
      If vForm.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then mvFindSinceSaveAsSS = False
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub MailingHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      AppHelper.ProcessSelectionSetMailing(MainHelper.MainForm, mvQBESelectionSet, False)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub SendEmailHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If EMailApplication.EmailInterface.CanEMail Then
        Dim vForm As frmSelectItems = New frmSelectItems(mvQBESelectionSet, True)
        If vForm.ShowDialog = DialogResult.OK Then
          EMailApplication.EmailInterface.SendMail(Me, EmailInterface.SendEmailOptions.seoAddressResolveUI Or EmailInterface.SendEmailOptions.seoMultipleRecipients, "", "", vForm.EMailAddresses)
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub GoToListManagerHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Dim vSelectionSet As Integer = mvQBESelectionSet
      If mvFindSinceSaveAsSS OrElse Not SelectionSetExists(vSelectionSet) Then
        Dim vList As New ParameterList(True)
        vList.Add("FromTemporaryTable", mvQBESelectionSet)
        If SelectionSetExists(vSelectionSet) Then vSelectionSet = 0
        Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vSelectionSet, Nothing, vList, Nothing)
        If vForm.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
          vSelectionSet = IntegerValue(vForm.ReturnList("SelectionSetNumber"))
          mvFindSinceSaveAsSS = False
        Else
          Exit Sub
        End If
      End If
      Dim vLM As New ListManager(vSelectionSet, False)
      vLM.Show()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub ReportMailmergeHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Dim vSelectionSet As Integer = mvQBESelectionSet
      Dim vRDS As New frmReportDataSelection(vSelectionSet, True)
      vRDS.Show()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Function SelectionSetExists(ByVal pSelectionSetNumber As Integer) As Boolean
    Dim vList As New ParameterList(True)
    vList("SelectionSet") = pSelectionSetNumber.ToString
    Dim vResult As ParameterList = DataHelper.GetLookupItem(CareNetServices.XMLLookupDataTypes.xldtSelectionSets, vList)
    If vResult.ContainsKey("SelectionSet") Then Return True
  End Function

  Private Sub mvBrowserMenu_Remove(ByVal pEntityType As CDBNETCL.HistoryEntityTypes, ByVal pNumber As Integer) Handles mvBrowserMenu.Remove
    dgrResults.DeleteRow(dgrResults.CurrentRow)
    SetResultMessages(0)
    DataHelper.DeleteSelectionSetContact(mvQBESelectionSet, pNumber, 0)
  End Sub
  Private Sub mvDocumentMenu_ShowRelatedDocument() Handles mvDocumentMenu.ShowRelatedDocument
    If mvDocumentNumber > 0 Then
      Dim vList As New ParameterList
      vList.IntegerValue("CommunicationsLogNumber1") = mvDocumentNumber
      vList("FinderCaption") = ControlText.FrmRelatedDocumentsFinder
      Dim vFinder As New frmFinder(CareNetServices.XMLDataFinderTypes.xdftDocuments, vList)
      vFinder.IsMainFinder = True
      MainHelper.SetMDIParent(vFinder)
      vFinder.Show()
    End If
  End Sub

  Public Property MultipleSelect As Boolean
    Get
      Return Me.dgrResults.MultipleSelect
    End Get
    Set(pAllow As Boolean)
      Me.dgrResults.MultipleSelect = pAllow
    End Set
  End Property

  Protected Overrides Sub ProcessStandingOrders()
    Try
      If mvSequenceNumber > 0 Then
        Dim vParams As New ParameterList(True)
        vParams.IntegerValue("NumberOfEntries") = mvNoOfItems
        vParams("BatchTotal") = mvTotalAmount.ToString("0.00")
        vParams("TTYNumber") = mvConnectionID
        Me.Close()
        FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtManualSOReconciliation, vParams, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysRun)
      End If
    Catch vEX As Exception
      If mvSequenceNumber > 0 Then
        'Manual SO Reconciliation errors
        ShowErrorMessage(vEX.Message)
      Else
        Throw vEX
      End If
    End Try
  End Sub

  Private Sub ViewMailingDocument()
    Try
      Dim vMailingNumber As Integer = IntegerValue(dgrResults.GetValue(dgrResults.CurrentRow, "MailingNumber"))
      If vMailingNumber > 0 Then
        Dim vParamList As New ParameterList(True)
        vParamList("MailingNumber") = vMailingNumber.ToString
        Dim vMailingHistoryDocumentCount As Integer = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMailingHistoryDocuments, vParamList)

        If IntegerValue(dgrResults.GetValue(dgrResults.CurrentRow, "EmailJobNumber")) > 0 Then
          Dim vStdDocument As String = "HTML"
          Dim vList As New ParameterList(True)
          vList("MailingNumber") = vMailingNumber.ToString
          Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtEmailJobs, vList)
          If vRow IsNot Nothing Then vStdDocument = vRow("StandardDocument").ToString
          Dim vSDList As New ParameterList(True)
          vSDList("StandardDocument") = vStdDocument
          vRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vSDList).Rows(0)
          Dim vApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocFileExtension").ToString)
          If vMailingHistoryDocumentCount = 0 Then vMailingNumber = 0
          vApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocFileExtension").ToString, "", False, False, True, vMailingNumber)
        Else
          Dim vApplication As ExternalApplication = GetDocumentApplication(".doc")
          vApplication.ViewMailingHistoryDocument(vMailingNumber, ".doc")
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub EditBatchDetails(ByVal sender As Object, ByVal pBatchInfo As BatchInfo, ByVal pTransactionNumber As Integer)
    Dim vTA As New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_edit_trans)), pBatchInfo.BatchNumber, False, pTransactionNumber)
    vTA.BatchNumber = pBatchInfo.BatchNumber
    vTA.TransactionNumber = pTransactionNumber
    vTA.BatchInfo = pBatchInfo
    vTA.BatchLocked = True
    Dim vList As New ParameterList()
    vList("EditFromBatchDetails") = "Y"
    FormHelper.RunTraderApplication(vTA, vList, Nothing, BatchInfo.AdjustmentTypes.None)
  End Sub

  Private Sub RefreshBatchDetails(ByVal sender As Object)
    RefreshData()
    Dim vList As New ParameterList(True)
    vList.IntegerValue("BatchNumber") = mvSelectedItem
    DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoUnlockBatch, vList)
  End Sub

  Private Sub dgr_EntitySelected(sender As Object, entityID As Integer, pType As HistoryEntityTypes) Handles dgr2.EntitySelected, dgr3.EntitySelected, dgr4.EntitySelected
    MainHelper.NavigateHistoryItem(pType, entityID)
  End Sub

  Private Sub dgr_ExamCentreSelected(sender As Object, pRow As Integer, pExamCentreID As Integer) Handles dgr2.ExamCentreSelected, dgr3.ExamCentreSelected, dgr4.ExamCentreSelected
    'call the standard method for navigation handling
    dgr_EntitySelected(sender, pExamCentreID, HistoryEntityTypes.hetExamCentres)
  End Sub
  Private Sub dgr_WorkstreamSelected(sender As Object, pRow As Integer, pWorkstreamID As Integer) Handles dgr2.WorkstreamSelected, dgr3.WorkstreamSelected, dgr4.WorkstreamSelected
    'call the standard method for navigation handling
    dgr_EntitySelected(sender, pWorkstreamID, HistoryEntityTypes.hetWorkstreams)
  End Sub
End Class
