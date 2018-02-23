Public Class frmCardMaintenance

  Protected mvParentForm As MaintenanceParentForm
  Protected mvMaintenanceType As CareServices.XMLMaintenanceControlTypes
  Protected mvSelectedRow As Integer
  Protected mvRefreshParent As Boolean
  Protected mvReturnList As ParameterList

  Private mvContactDataType As CareServices.XMLContactDataSelectionTypes
  Private mvContactInfo As ContactInfo
  Private mvRelatedContact As ContactInfo
  Private mvSelectionSetNumber As Integer
  Private mvList As ParameterList
  Private mvEMailMessage As EMailMessage
  Private mvOurReference As String
  Private mvDocumentFile As String
  Private mvDocumentSaved As Boolean
  Private mvPackage As String
  Private mvExtension As String
  Private mvDefaultSize As New Size(800, 600)
  Private mvOrganisationNumber As Integer
  Private mvContactAddressNumber As Integer
  Private mvSetRelatedContact As ContactInfo
  Private mvCommsLogSaved As Boolean

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub
  Public Sub New(ByVal pContactInfo As ContactInfo, ByVal pList As ParameterList, ByVal pMDI As Boolean)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pContactInfo, pList, pMDI)
  End Sub
  Public Sub New(ByVal pForm As MaintenanceParentForm, ByVal pContactInfo As ContactInfo, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pDataSet As DataSet, ByVal pEdit As Boolean, ByVal pRow As Integer, Optional ByVal pMaintenanceType As CareServices.XMLMaintenanceControlTypes = CareServices.XMLMaintenanceControlTypes.xmctNone, Optional ByVal pList As ParameterList = Nothing)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pForm, pContactInfo, pType, pDataSet, pEdit, pRow, pMaintenanceType, pList)
  End Sub
  Public Sub New(ByVal pMaintenanceType As CareServices.XMLMaintenanceControlTypes, ByVal pNumber As Integer, Optional ByVal pParentForm As MaintenanceParentForm = Nothing, Optional ByVal pList As ParameterList = Nothing)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pMaintenanceType, pNumber, pParentForm, pList)
  End Sub

#Region "Maintenance Parent Form Methods"

  Public Overrides ReadOnly Property SizeMaintenanceForm() As Boolean
    Get
      Return True
    End Get
  End Property

  Public Overrides Sub RefreshData(ByVal pType As CareServices.XMLMaintenanceControlTypes)
    'Do nothing
  End Sub

#End Region

  Private Sub InitialiseControls()
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me, "frmCardMaintenance")     'Pass in class name as could be for inherited class
#End If
    SetControlColors()
  End Sub
  Private Sub InitialiseControls(ByVal pContactInfo As ContactInfo, ByVal pList As ParameterList, ByVal pMDI As Boolean)
    'Only called for New Contact or Organisation
    Me.SuspendLayout()
    SetControlColors()
    splTop.Panel1Collapsed = True         'No Header
    splMaint.Panel1Collapsed = True       'No Combo
    splBottom.Panel1Collapsed = True      'No Treeview
    If pMDI Then Me.MdiParent = MDIForm
    mvParentForm = Nothing
    mvContactInfo = pContactInfo
    mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation
    mvMaintenanceType = GetMaintenanceType(mvContactDataType, pContactInfo.ContactType)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctContact
        If AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_contact_deduplication) = "I" Then mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry
      Case CareServices.XMLMaintenanceControlTypes.xmctOrganisation
        If AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_contact_deduplication) = "I" Then mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry
    End Select
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry Then
      pList("Country") = AppValues.DefaultCountryCode
    End If
    mvSelectedRow = -1
    cmdDelete.Visible = False
    cmdNew.Visible = False
    cmdDefault.Visible = False
    HideGrid()
    epl.Init(New EditPanelInfo(mvMaintenanceType, mvContactInfo, , mvContactDataType))
    Dim vSetOrganisation As Boolean = False
    If pList.ContainsKey("CreateAtOrganisationNumber") Then
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry Then
        Dim vList As New ParameterList(True)
        mvOrganisationNumber = CInt(pList("CreateAtOrganisationNumber"))
        Dim vRow As DataRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, mvOrganisationNumber)
        If vRow IsNot Nothing Then
          pList("Name") = vRow("ContactName").ToString
          mvContactAddressNumber = CInt(vRow("AddressNumber"))
          vSetOrganisation = True
        End If
      Else
        pList("OrganisationNumber") = pList("CreateAtOrganisationNumber")
      End If
    End If
    If mvContactInfo.CreateAtAddressNumber > 0 Then mvContactAddressNumber = mvContactInfo.CreateAtAddressNumber
    epl.Populate(pList)
    If vSetOrganisation Then UseAddressNumber()
    If epl.Caption.Length > 0 Then Me.Text = epl.Caption
    Me.ClientSize = New Size(Me.ClientSize.Width, epl.RequiredHeight + bpl.Height)
    'DebugPrint("End of Initialise Controls")
    Me.ResumeLayout()
  End Sub
  Private Sub InitialiseControls(ByVal pMaintenanceType As CareServices.XMLMaintenanceControlTypes, ByVal pNumber As Integer, Optional ByVal pParentForm As MaintenanceParentForm = Nothing, Optional ByVal pList As ParameterList = Nothing)
    'Called for Document Maintenance and possibly others e.g. Action maintenance
    SetControlColors()
    splTop.Panel1Collapsed = True         'No Header
    splMaint.Panel1Collapsed = True       'No Combo
    splBottom.Panel1Collapsed = True      'No Treeview
    Select Case pMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctCriterialSet, CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, CareServices.XMLMaintenanceControlTypes.xmctMailingOptions, CareServices.XMLMaintenanceControlTypes.xmctBatches
        'Dialog
      Case Else
        Me.MdiParent = MDIForm
    End Select
    mvParentForm = pParentForm
    If Not mvParentForm Is Nothing Then
      If mvParentForm.SizeMaintenanceForm Then Me.Size = mvParentForm.Size
      mvContactInfo = mvParentForm.ContactInfo
    Else
      mvContactInfo = New ContactInfo(ContactInfo.ContactTypes.ctContact, "")
    End If
    mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtNone
    mvMaintenanceType = pMaintenanceType
    mvList = pList
    cmdDefault.Visible = False
    cmdNew.Visible = False
    HideGrid()
    Select Case pMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctDocument, CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
        mvContactInfo.SelectedDocumentNumber = pNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        mvContactInfo.SelectedActionNumber = pNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctActionSchedule
        mvContactInfo.SelectedActionNumber = pNumber
        cmdDelete.Visible = False
      Case CareServices.XMLMaintenanceControlTypes.xmctSelectionSet
        mvSelectionSetNumber = pNumber
    End Select
    SetCommandButtons()
    epl.Init(New EditPanelInfo(mvMaintenanceType, mvContactInfo))
    If epl.Caption.Length > 0 Then Me.Text = epl.Caption
    Me.ClientSize = New Size(Me.ClientSize.Width, epl.RequiredHeight + bpl.Height)
    If pMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctBatches Then
      epl.Populate(pList)
    End If
  End Sub
  Private Sub InitialiseControls(ByVal pForm As MaintenanceParentForm, ByVal pContactInfo As ContactInfo, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pDataSet As DataSet, ByVal pEdit As Boolean, ByVal pRow As Integer, ByVal pMaintenanceType As CareServices.XMLMaintenanceControlTypes, Optional ByVal pList As ParameterList = Nothing)
    SetControlColors()
    splTop.Panel1Collapsed = True         'No Header
    splMaint.Panel1Collapsed = True       'No Combo
    splBottom.Panel1Collapsed = True      'No Treeview
    Me.MdiParent = MDIForm
    mvParentForm = pForm
    If Not mvParentForm Is Nothing AndAlso mvParentForm.SizeMaintenanceForm Then Me.Size = mvParentForm.Size
    mvContactInfo = pContactInfo
    mvContactDataType = pType
    mvMaintenanceType = pMaintenanceType
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone Then
      mvMaintenanceType = GetMaintenanceType(mvContactDataType, pContactInfo.ContactType)
    End If
    mvList = pList
    cmdDelete.Visible = CanDelete(mvMaintenanceType)
    cmdDefault.Visible = (mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses Or mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers)
    cmdNew.Visible = (mvContactDataType <> CareServices.XMLContactDataSelectionTypes.xcdtContactNotes)
    epl.Init(New EditPanelInfo(mvMaintenanceType, mvContactInfo))
    If epl.Caption.Length > 0 Then
      Select Case mvContactDataType
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom
          Me.Text = String.Format(InformationMessages.imLinkMaintenanceCaptionFrom, epl.Caption, mvContactInfo.ContactName)
        Case CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo
          Me.Text = String.Format(InformationMessages.imLinkMaintenanceCaptionTo, epl.Caption, mvContactInfo.ContactName)
        Case Else
          Me.Text = epl.Caption
      End Select
    End If
    SetCommandButtons()
    If mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation Then
      cmdNew.Visible = False
      HideGrid()
    Else
      If pEdit Then
        mvSelectedRow = pRow
      Else
        mvSelectedRow = -1
      End If
      dgr.AutoSetHeight = True
      dgr.Populate(pDataSet, , mvSelectedRow)
      splRight.SplitterDistance = dgr.RequiredHeight
      If dgr.RowCount = 0 Then mvSelectedRow = -1
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, CareServices.XMLMaintenanceControlTypes.xmctActionLink
          dgr.AllowDrop = True
      End Select
    End If
  End Sub

  Private Sub SetControlColors()
    Me.BackColor = DisplayTheme.FormBackColor
    bpl.UseTheme()
  End Sub
  Private Sub SetCommandButtons()
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        cmdLink1.Text = ControlText.cmdActionLinks
        cmdLink1.Visible = True
        cmdLink2.Text = ControlText.cmdActionSubjects
        cmdLink2.Visible = True
      Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
        cmdLink1.Text = ControlText.cmdAddressUsages
        cmdLink1.Visible = True
        If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
          cmdLink2.Text = ControlText.cmdCloseSite
          cmdLink2.Visible = True
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctDocument, CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
        bpl.DockingStyle = ButtonPanel.ButtonPanelDockingStyles.bpdsRight
        cmdLink1.Text = ControlText.cmdDocumentLinks
        cmdLink1.Visible = True
        cmdLink2.Text = ControlText.cmdDocumentSubjects
        cmdLink2.Visible = True
        cmdCreateOrEdit.Text = ControlText.cmdDocumentCreate
        cmdOther.Text = ControlText.cmdEMail
        cmdReply.Text = ControlText.cmdDocumentReply
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocument Then
          cmdCreateOrEdit.Visible = True
          cmdCreateOrEdit.Enabled = False
          cmdOther.Visible = True
          cmdOther.Enabled = EMailApplication.EmailInterface.CanEMail
          cmdPrint.Visible = True
          cmdPrint.Enabled = False
        End If
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocument Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEMailDocument Then
          If mvList IsNot Nothing AndAlso mvList.Contains("AsReply") AndAlso mvList("AsReply") = "Y" Then
            cmdReply.Visible = False
          Else
            cmdReply.Visible = True
          End If
        End If
    End Select
  End Sub
  Protected Sub HideGrid()
    dgr.Visible = False
    splRight.Panel1Collapsed = True
  End Sub

  Private Sub frmCardMaintenance_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    epl.ClearDataSources(epl)
  End Sub

  Protected Overridable Sub frmCardMaintenance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'DebugPrint("Card Maintenance Load")
    If Me.DesignMode Then Return
    Me.SuspendLayout()
    If mvParentForm IsNot Nothing AndAlso mvParentForm.SizeMaintenanceForm Then
      If mvParentForm.MdiParent IsNot Nothing And Me.MdiParent Is Nothing Then
        'We are putting up a dialog type form on top of an mdi child maintenance form e.g. contact maintenance
        'So we need to make the location relative to the screen
        Location = mvParentForm.PointToScreen(mvParentForm.Location)
      Else
        Location = mvParentForm.Location
      End If
      Size = mvParentForm.Size              'Required here for Windows 2000
      mvParentForm.Enabled = False
    Else
      If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry OrElse mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry Then
        Width = mvDefaultSize.Width
      Else
        Size = mvDefaultSize
      End If
      If Not mvParentForm Is Nothing Then mvParentForm.Enabled = False
      If MdiParent Is Nothing Then
        Location = MDIForm.PointToScreen(MDILocation(mvDefaultSize.Width, mvDefaultSize.Height))
      Else
        Location = MDILocation(mvDefaultSize.Width, mvDefaultSize.Height)
      End If
      'DebugPrint("Card Maintenance Set Size and Location")
    End If
    If mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation Then
      If mvContactInfo.ContactNumber > 0 Then
        epl.Populate(DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, mvContactInfo.ContactNumber))
        If Not AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciContactChangeDepartment) Then epl.EnableControl("Department", False)
        If epl.GetValue("Status").Length > 0 And Not AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciContactChangeStatus) Then
          epl.EnableControl("Status", False)
          epl.EnableControl("StatusDate", False)
          epl.EnableControl("StatusReason", False)
        Else
          epl.FindTextLookupBox("Status").FillComboWithRestriction(epl.GetValue("Status"), mvContactInfo.ContactGroup)
        End If
        If Not AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciContactChangeSource) Then
          epl.EnableControl("Source", False)
          epl.EnableControl("SourceDate", False)
        End If
        If Not AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciContactChangeOwnershipDetails) Then
          Dim vPrincipalDepartment As String = epl.FindTextLookupBox("OwnershipGroup").GetDataRowItem("PrincipalDepartment")
          If vPrincipalDepartment.Length > 0 AndAlso vPrincipalDepartment <> DataHelper.UserInfo.Department Then
            epl.EnableControl("OwnershipGroup", False)
            epl.EnableControl("PrincipalUser", False)
            epl.EnableControl("PrincipalUserReason", False)
          End If
        End If
      Else
        'New contact - may have a postcode and if so should popup address finder
        'If adding a new contact at an existing organisation then postcode will not be present
        Dim vControl As Control = FindControl(epl, "Postcode", False)
        If vControl IsNot Nothing Then epl.PopupAddressCheck(vControl)
        SetDefaults()
        epl.SetDependantValues()
      End If
      epl.SetDependancies("Status")           'Set Status Date and Status Reason enabled
      epl.SetDependancies("VatCategory")      'Set VAT number enabled
      epl.SetDependancies("PrincipalUser")    'Set Principal User Reason enabled
    ElseIf mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtNone Then
      If mvContactInfo IsNot Nothing AndAlso mvContactInfo.SelectedDocumentNumber > 0 Then
        Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentInformation, mvContactInfo.SelectedDocumentNumber))
        epl.Populate(vDataRow)
        Dim vIsReply As Boolean = mvList IsNot Nothing AndAlso mvList.Contains("AsReply") AndAlso mvList("AsReply") = "Y"
        mvExtension = vDataRow("ExternalApplicationExtension").ToString
        epl.SetDocumentStyle(epl.FindComboBox("DocumentStyle"))
        epl.SetDocumentType(epl.FindComboBox("Package"))
        epl.EnableControl("DocumentStyle", vIsReply)
        epl.SetValue("StandardDocument", vDataRow("StandardDocument").ToString, Not vIsReply)
        epl.SetValue("Package", vDataRow("ExternalApplicationCode").ToString, Not vIsReply)
        epl.SetValue("DocumentType", vDataRow("DocumentType").ToString, Not vIsReply)
        Dim vEnabledItems() As String = {"Topic", "SubTopic", "DocumentSubject", "Precis"}
        For Each vItem As String In vEnabledItems
          epl.SetValue(vItem, MultiLine(vDataRow(vItem).ToString))
        Next
        epl.DataChanged = False
        If vIsReply Then epl.SetValue("Direction", "O", True)
        cmdCreateOrEdit.Text = ControlText.cmdDocumentEdit
        cmdReply.Enabled = cmdReply.Visible AndAlso epl.GetValue("Direction") = "I"
        'TODO Change to use DocumentName
        UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetDocuments, mvContactInfo.SelectedDocumentNumber, mvContactInfo.SelectedDocumentNumber & " - " & vDataRow.Item("OurReference").ToString)
      ElseIf mvContactInfo IsNot Nothing AndAlso mvContactInfo.SelectedActionNumber > 0 Then
        Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetActionData(CareServices.XMLActionDataSelectionTypes.xadtActionInformation, mvContactInfo.SelectedActionNumber))
        epl.Populate(vDataRow)
        UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetActions, mvContactInfo.SelectedActionNumber, vDataRow.Item("ActionDesc").ToString)
      Else
        SetDefaults()
        mvSelectedRow = -1
        cmdDelete.Visible = False
      End If
    Else
      If mvSelectedRow >= 0 Then
        SelectRow(mvSelectedRow)
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionLink Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocumentLink Then SetDefaults()
      Else
        SetDefaults()
        SetCommandsForNew()
      End If
    End If
    'DebugPrint("Card Maintenance About to reposition buttons")
    bpl.RepositionButtons()
    epl.SetFocus()
    'DebugPrint("Card Maintenance End of Load")
    Me.ResumeLayout()
  End Sub
  Private Sub frmCardMaintenance_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SystemColorsChanged
    DisplayTheme.UpdateThemeColors()
    SetControlColors()
  End Sub
  Private Sub frmCardMaintenance_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    Try
      e.Cancel = False        'Sometimes comes in as true???
      If epl.DataChanged Then
        If ConfirmCancel() = False Then e.Cancel = True
      End If
      If Not e.Cancel Then
        If Not mvDocumentFile Is Nothing Then
          Dim vFileInfo As New FileInfo(mvDocumentFile)
          If vFileInfo.Exists Then vFileInfo.Delete()
        End If
        If mvList IsNot Nothing AndAlso mvList.Contains("AsReply") AndAlso mvList("AsReply") = "Y" And Not mvCommsLogSaved Then
          Dim vList As New ParameterList(True)
          vList.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
          DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctDocument, vList)
          UserHistory.RemoveOtherHistoryNode(HistoryEntityTypes.hetDocuments, mvContactInfo.SelectedDocumentNumber)
        End If
        If Not mvParentForm Is Nothing Then
          mvParentForm.Enabled = True
          mvParentForm.BringToFront()
          If mvRefreshParent Then mvParentForm.RefreshData(mvMaintenanceType)
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overridable Sub SelectRow(ByVal pRow As Integer)
    Try
      If pRow >= 0 Then
        Dim vList As New ParameterList(True)
        Dim vCount As Integer = vList.Count
        GetPrimaryKeyValues(vList, pRow, False)
        If vList.Count > vCount Then
          Dim vDataRow As DataRow
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctAction
              vDataRow = DataHelper.GetRowFromDataSet(DataHelper.GetActionData(CareServices.XMLActionDataSelectionTypes.xadtActionInformation, vList.IntegerValue("ActionNumber")))
              epl.Populate(vDataRow)
            Case CareServices.XMLMaintenanceControlTypes.xmctActionLink, CareServices.XMLMaintenanceControlTypes.xmctActionTopic
              'Don't allow edit
            Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink
              'Don't allow edit
              Select Case vList("DocumentLinkType")
                Case "A", "S"
                  cmdDelete.Enabled = False
                Case Else
                  cmdDelete.Enabled = True
              End Select
            Case CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
              'Don't allow edit
              cmdDelete.Enabled = dgr.GetValue(pRow, "Primary") <> "Y"
            Case Else
              vDataRow = DataHelper.GetContactItem(mvContactDataType, mvContactInfo.ContactNumber, vList)
              epl.Populate(vDataRow)
              cmdDelete.Enabled = True
              Select Case mvMaintenanceType
                Case CareServices.XMLMaintenanceControlTypes.xmctNumber
                  epl.SetDependancies("Device")
                Case CareServices.XMLMaintenanceControlTypes.xmctLink
                  'If mvList IsNot Nothing AndAlso mvList.IntegerValue("ContactNumber2") > 0 Then     'We are restricting to links to this other contact
                  epl.EnableControl("ContactGroup", False)
                  epl.EnableControl("ContactNumber2", False)                      'Don't allow user to change contact the link is with
                Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
                  If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
                    cmdLink2.Enabled = Not vDataRow("Historical").ToString.StartsWith("Y") AndAlso Not vDataRow("Default").ToString.StartsWith("Y")
                  End If
                Case CareServices.XMLMaintenanceControlTypes.xmctSuppression
                  epl.EnableControl("Suppression", False)
              End Select
          End Select
        End If
        mvSelectedRow = pRow
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub GetPrimaryKeyValues(ByVal pList As ParameterList, ByVal pRow As Integer, ByVal pForUpdate As Boolean)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAction, CareServices.XMLMaintenanceControlTypes.xmctActionSchedule
        If dgr.Visible Then
          pList("ActionNumber") = dgr.GetValue(pRow, "ActionNumber")
        Else
          pList.IntegerValue("ActionNumber") = mvContactInfo.SelectedActionNumber
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
        pList("AddressNumber") = dgr.GetValue(pRow, "AddressNumber")
        mvContactInfo.SelectedAddressNumber = pList.IntegerValue("AddressNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink
        pList.IntegerValue("ActionNumber") = mvContactInfo.SelectedActionNumber
        pList("ContactNumber") = dgr.GetValue(pRow, "ContactNumber")
        pList("ActionLinkType") = dgr.GetValue(pRow, "LinkType")
      Case CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        pList.IntegerValue("ActionNumber") = mvContactInfo.SelectedActionNumber
        pList("Topic") = dgr.GetValue(pRow, "TopicCode")
        pList("SubTopic") = dgr.GetValue(pRow, "SubTopicCode")
      Case CareServices.XMLMaintenanceControlTypes.xmctActivities, CareServices.XMLMaintenanceControlTypes.xmctPositionActivity
        If pForUpdate Then
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctPositionActivity Then
            pList("ContactPositionNumber") = dgr.GetValue(pRow, "ContactPositionNumber")
          Else
            pList.IntegerValue("OldContactNumber") = mvContactInfo.ContactNumber
          End If
          pList("OldActivity") = dgr.GetValue(pRow, "ActivityCode")
          pList("OldActivityValue") = dgr.GetValue(pRow, "ActivityValueCode")
          pList("OldSource") = dgr.GetValue(pRow, "SourceCode")
          pList("OldValidFrom") = dgr.GetValue(pRow, "ValidFrom")
          pList("OldValidTo") = dgr.GetValue(pRow, "ValidTo")
        Else
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctPositionActivity Then
            pList("ContactPositionNumber") = dgr.GetValue(pRow, "ContactPositionNumber")
          End If
          pList("Activity") = dgr.GetValue(pRow, "ActivityCode")
          pList("ActivityValue") = dgr.GetValue(pRow, "ActivityValueCode")
          pList("Source") = dgr.GetValue(pRow, "SourceCode")
          pList("ValidFrom") = dgr.GetValue(pRow, "ValidFrom")
          pList("ValidTo") = dgr.GetValue(pRow, "ValidTo")
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink
        pList.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
        If dgr.GetValue(pRow, "AddressNumber").Length > 0 Then
          pList("ContactNumber") = dgr.GetValue(pRow, "ContactNumber")
        Else
          pList("DocumentNumber2") = dgr.GetValue(pRow, "ContactNumber")
        End If
        pList("DocumentLinkType") = dgr.GetValue(pRow, "LinkType")
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
        pList.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
        pList("Topic") = dgr.GetValue(pRow, "TopicCode")
        pList("SubTopic") = dgr.GetValue(pRow, "SubTopicCode")
      Case CareServices.XMLMaintenanceControlTypes.xmctDocument, CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
        If mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtNone Then
          pList.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
        Else
          pList("DocumentNumber") = dgr.GetValue(pRow, "DocumentNumber")
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctLink
        If pForUpdate Then
          pList("OldRelationship") = dgr.GetValue(pRow, "RelationshipCode")
        Else
          pList("Relationship") = dgr.GetValue(pRow, "RelationshipCode")
          pList("ContactNumber2") = dgr.GetValue(pRow, "ContactNumber")
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctNumber
        If pForUpdate Then pList.IntegerValue("OldContactNumber") = mvContactInfo.ContactNumber
        pList("CommunicationNumber") = dgr.GetValue(pRow, "CommunicationNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctPosition
        pList("ContactPositionNumber") = dgr.GetValue(pRow, "ContactPositionNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctReference
        If pForUpdate Then
          pList("OldDataSource") = dgr.GetValue(pRow, "DataSource")
          pList("OldExternalReference") = dgr.GetValue(pRow, "ExternalReference")
        Else
          pList("DataSource") = dgr.GetValue(pRow, "DataSource")
          pList("ExternalReference") = dgr.GetValue(pRow, "ExternalReference")
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctRole
        pList("ContactRoleNumber") = dgr.GetValue(pRow, "ContactRoleNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctSuppression
        pList("Suppression") = dgr.GetValue(pRow, "SuppressionCode")
      Case CareServices.XMLMaintenanceControlTypes.xmctStickyNote
        pList("NoteNumber") = dgr.GetValue(pRow, "NoteNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctDepartmentNotes
        pList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber          'Will get added later but gets this working
      Case CareServices.XMLMaintenanceControlTypes.xmctAddressUsage
        pList.IntegerValue("AddressNumber") = mvContactInfo.SelectedAddressNumber
        pList("AddressUsage") = dgr.GetValue(pRow, "AddressUsage")
    End Select
  End Sub
  Private Sub GetAdditionalKeyValues(ByVal pList As ParameterList)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctSelectionSet
        pList.IntegerValue("SelectionSetNumber") = mvSelectionSetNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctAddressUsage
        pList.IntegerValue("AddressNumber") = mvContactInfo.SelectedAddressNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctPosition
        If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
          pList.IntegerValue("OrganisationNumber") = mvContactInfo.ContactNumber
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctRole
        If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
          pList.IntegerValue("OrganisationNumber") = mvContactInfo.ContactNumber
          pList.IntegerValue("ContactNumber2") = mvContactInfo.SelectedContactNumber2
        Else
          pList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
          pList.IntegerValue("OrganisationNumber") = mvContactInfo.SelectedContactNumber2
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctPositionActivity
        pList.IntegerValue("ContactPositionNumber") = mvContactInfo.SelectedContactPositionNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctDocument
        pList("DocumentNumber") = epl.GetValue("DocumentNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
        pList.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink, CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        pList.IntegerValue("ActionNumber") = mvContactInfo.SelectedActionNumber
      Case CareServices.XMLMaintenanceControlTypes.xmctTCRDocument
        If epl.GetValue("Direction") = "I" Then
          pList("DocumentClass") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_in_document_class)
          pList("DocumentType") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_in_document_type)
        Else
          pList("DocumentClass") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_document_class)
          pList("DocumentType") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_document_type)
        End If
        pList("Dated") = Date.Now.ToShortDateString
        pList("DocumentNumber") = epl.GetValue("DocumentNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
        pList("DocumentClass") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.email_in_document_class)
        pList("DocumentType") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.email_in_document_type)
        pList("Dated") = mvEMailMessage.DateReceived
        'pList("Dated") = Date.Now.ToShortDateString
        pList("Direction") = "I"
        pList("DocumentNumber") = epl.GetValue("DocumentNumber")
      Case CareServices.XMLMaintenanceControlTypes.xmctContact, CareServices.XMLMaintenanceControlTypes.xmctContactEntry
        If mvContactAddressNumber > 0 Then pList.IntegerValue("AddressNumber") = mvContactAddressNumber
        If mvOrganisationNumber > 0 Then pList.IntegerValue("OrganisationNumber") = mvOrganisationNumber
    End Select
  End Sub
  Private Function NoEditingAllowed() As Boolean
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink, CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        Return True
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
        Return True
    End Select
  End Function
  Private Function EditingExistingRecord() As Boolean
    If mvSelectedRow >= 0 AndAlso NoEditingAllowed() = False Then Return True
  End Function

  Protected Overridable Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub
  Protected Overridable Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click, cmdDefault.Click
    Try
      'TODO Confirm Update or Insert 
      Dim vDefault As Boolean
      If (mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses _
        Or mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers) _
        AndAlso sender Is cmdDefault Then vDefault = True
      If ProcessSave(vDefault) Then
        If dgr.Visible = False Then     'Only one record
          Me.DialogResult = Windows.Forms.DialogResult.OK
          Me.Close()
        Else
          RePopulateGrid()
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    'TODO Confirm cancel changes
    'Clear selection on display grid and set defaults for new record
    Try
      ProcessNew()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Enum CmdLinkTypes
    cltNone
    cltLinksUsages
    cltAnalysisCloseSite
    cltEMail
    cltReply
    cltPrint
  End Enum

  Protected Overridable Sub cmdLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLink1.Click, cmdLink2.Click, cmdOther.Click, cmdReply.Click, cmdPrint.Click
    Dim vAction As CmdLinkTypes
    Dim vForm As frmCardMaintenance = Nothing
    Dim vDataSet As DataSet
    Dim vProceed As Boolean = True

    Dim vCursor As New BusyCursor
    Try
      If epl.DataChanged OrElse Not EditingExistingRecord() Then vProceed = ProcessSave(False)
      If vProceed Then
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctAddresses Then
          RePopulateGrid()
          mvSelectedRow = dgr.FindRow("AddressNumber", mvContactInfo.SelectedAddressNumber.ToString)
          dgr.SelectRow(mvSelectedRow)
        Else
          mvSelectedRow = 0 'Make the form think that we are now dealing with an existing record
        End If
        If sender Is cmdLink1 Then
          vAction = CmdLinkTypes.cltLinksUsages
        ElseIf sender Is cmdLink2 Then
          vAction = CmdLinkTypes.cltAnalysisCloseSite
        ElseIf sender Is cmdOther Then
          vAction = CmdLinkTypes.cltEMail
        ElseIf sender Is cmdPrint Then
          vAction = CmdLinkTypes.cltPrint
        Else
          vAction = CmdLinkTypes.cltReply
        End If
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctAction
            If vAction = CmdLinkTypes.cltLinksUsages Then
              vDataSet = DataHelper.GetActionData(CareServices.XMLActionDataSelectionTypes.xadtActionLinks, mvContactInfo.SelectedActionNumber)
              vForm = New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtNone, vDataSet, False, 0, CareServices.XMLMaintenanceControlTypes.xmctActionLink)
            Else
              vDataSet = DataHelper.GetActionData(CareServices.XMLActionDataSelectionTypes.xadtActionSubjects, mvContactInfo.SelectedActionNumber)
              vForm = New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtNone, vDataSet, False, 0, CareServices.XMLMaintenanceControlTypes.xmctActionTopic)
            End If
          Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
            If vAction = CmdLinkTypes.cltLinksUsages Then           'Usages
              Dim vList As New ParameterList(True)
              vList.IntegerValue("AddressNumber") = mvContactInfo.SelectedAddressNumber
              vDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressUsages, mvContactInfo.ContactNumber, vList)
              vForm = New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactAddressUsages, vDataSet, False, 0, CareServices.XMLMaintenanceControlTypes.xmctAddressUsage)
            Else                        'Close Site
              Dim vOldAddressNumber As Integer = CInt(dgr.GetValue(dgr.CurrentRow, "AddressNumber"))
              vDataSet = DataHelper.GetContactData(mvContactDataType, mvContactInfo.ContactNumber)
              DataHelper.RemoveRows(vDataSet, "Historical", "Yes")
              DataHelper.RemoveRows(vDataSet, "AddressNumber", vOldAddressNumber.ToString)
              Dim vSA As frmSelectAddress = New frmSelectAddress(vDataSet, frmSelectAddress.SelectAddressTypes.satClosingSite)
              If vSA.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                mvRefreshParent = True
                DataHelper.CloseSite(mvContactInfo.ContactNumber, vOldAddressNumber, vSA.AddressNumber)
                RePopulateGrid()
              End If
            End If
          Case CareServices.XMLMaintenanceControlTypes.xmctDocument, CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
            Select Case vAction
              Case CmdLinkTypes.cltLinksUsages
                vDataSet = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentLinks, mvContactInfo.SelectedDocumentNumber)
                vForm = New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments, vDataSet, False, 0, CareServices.XMLMaintenanceControlTypes.xmctDocumentLink)
              Case CmdLinkTypes.cltAnalysisCloseSite
                vDataSet = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentSubjects, mvContactInfo.SelectedDocumentNumber)
                vForm = New frmCardMaintenance(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactDocuments, vDataSet, False, 0, CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic)
              Case CmdLinkTypes.cltEMail      'EMail
                vDataSet = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtRelatedDocuments, mvContactInfo.SelectedDocumentNumber)
                Dim vResult As DialogResult = Windows.Forms.DialogResult.OK
                Dim vTable As DataTable = Nothing
                If vDataSet.Tables.Contains("DataRow") Then
                  vTable = vDataSet.Tables("DataRow")
                  Dim vSelectForm As New frmSelectItems(vDataSet)
                  vResult = vSelectForm.ShowDialog(Me)
                Else
                  vTable = New DataTable("DataRow")
                  vTable.Columns.AddRange(New DataColumn() _
                  { _
                    New DataColumn("Select"), _
                    New DataColumn("DocumentNumber"), _
                    New DataColumn("OurReference"), _
                    New DataColumn("Subject"), _
                    New DataColumn("DocumentSource"), _
                    New DataColumn("WordProcessorDocument") _
                  })
                End If
                If vResult = Windows.Forms.DialogResult.OK Then
                  Dim vComboBox As ComboBox = epl.FindComboBox("DocumentStyle")
                  Dim vStyle As DataHelper.DocumentStyles = CType(vComboBox.SelectedValue, DataHelper.DocumentStyles)
                  Select Case vStyle
                    Case DataHelper.DocumentStyles.dsnPrecisOnly, DataHelper.DocumentStyles.dsnStandardDocumentPrecis
                      '
                    Case Else
                      Dim vRow As DataRow = vTable.NewRow
                      vRow.Item("Select") = "True"
                      vRow.Item("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
                      vRow.Item("OurReference") = epl.GetValue("OurReference")
                      vRow.Item("Subject") = epl.GetValue("DocumentSubject")
                      Dim vType As TextLookupBox = epl.FindTextLookupBox("DocumentType")
                      vRow.Item("DocumentSource") = vType.GetDataRowItem("DocumentSource")
                      vRow.Item("WordProcessorDocument") = vType.GetDataRowItem("WordProcessorDocument")
                      vTable.Rows.Add(vRow)
                  End Select
                  EMailApplication.SendDocumentAsEMail(Me, mvContactInfo.SelectedDocumentNumber, vTable)
                End If
              Case CmdLinkTypes.cltPrint
                Dim vApplication As ExternalApplication
                vApplication = GetApplication(False)
                If Not vApplication Is Nothing Then
                  vApplication.PrintDocument(mvContactInfo.SelectedDocumentNumber, mvExtension)
                  mvRefreshParent = True
                End If
              Case CmdLinkTypes.cltReply  'Reply
                'We've created/updated the document so now we'll generate a copy of that document with the Sender and Addressee reversed, and display that document to the user
                Dim vList As New ParameterList(True)
                vList("AsReply") = "Y"
                vList.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
                vList("Dated") = Date.Now.ToShortDateString
                mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
                vForm = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctDocument, mvReturnList.IntegerValue("DocumentNumber"), mvParentForm, vList)
                Me.Close()
            End Select
          Case Else
            vForm = Nothing
        End Select
        If Not vForm Is Nothing Then vForm.Show()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub cmdCreateOrEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateOrEdit.Click
    Dim vCursor As New BusyCursor
    Try
      Dim vApplication As ExternalApplication
      epl.DataChanged = True
      If EditingExistingRecord() Then
        vApplication = GetApplication(True)
        If Not vApplication Is Nothing Then
          vApplication.EditDocument(mvContactInfo.SelectedDocumentNumber, mvExtension)
          'cmdCreateOrEdit.Enabled = True
          'If we allow edit again it will get another copy of the document...
        End If
      Else
        Dim vList As New ParameterList
        If epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll) Then
          Dim vComboBox As ComboBox = epl.FindComboBox("DocumentStyle")
          Dim vStyle As DataHelper.DocumentStyles = CType(vComboBox.SelectedValue, DataHelper.DocumentStyles)
          Select Case vStyle
            Case DataHelper.DocumentStyles.dsnBlankDocument
              vApplication = GetApplication(True)
              If Not vApplication Is Nothing Then vApplication.EditNewDocument(Nothing, mvExtension)

            Case DataHelper.DocumentStyles.dsnTopAndTailedDocument
              Dim vRow As DataRow
              Dim vContactList As New ParameterList(True)
              vContactList.IntegerValue("ContactNumber") = vList.IntegerValue("AddresseeContactNumber")
              vContactList.IntegerValue("AddressNumber") = vList.IntegerValue("AddresseeAddressNumber")
              vRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, 0, vContactList)
              vList("Addressee") = vRow.Item("ContactName").ToString
              vList("AddresseeAddress") = vRow.Item("AddressMultiLine").ToString
              vList("Salutation") = vRow.Item("Salutation").ToString
              Dim vSender As Integer = vList.IntegerValue("SenderContactNumber")
              vRow = DataHelper.GetContactItem(CareServices.XMLContactDataSelectionTypes.xcdtContactUserInformation, vSender)
              vList("SignatureName") = vRow.Item("FullName").ToString
              vList("SignaturePosition") = vRow.Item("Position").ToString
              vApplication = GetApplication(True)
              If Not vApplication Is Nothing Then vApplication.EditNewDocument(vList, mvExtension)

            Case DataHelper.DocumentStyles.dsnStandardDocumentWithMerge, DataHelper.DocumentStyles.dsnStandardDocumentTemplate
              vApplication = GetApplication(True)
              If mvSetRelatedContact IsNot Nothing Then vList.IntegerValue("RelatedContactNumber") = mvSetRelatedContact.ContactNumber
              If Not vApplication Is Nothing Then vApplication.EditNewStandardDocument(vList, mvExtension, vStyle = DataHelper.DocumentStyles.dsnStandardDocumentWithMerge)
          End Select
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Protected Overridable Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      'TODO Confirm cancel changes and Confirm Delete
      Dim vList As New ParameterList(True)
      If mvContactInfo.ContactNumber > 0 Then vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
      GetPrimaryKeyValues(vList, mvSelectedRow, False)
      DataHelper.DeleteItem(mvMaintenanceType, vList)
      Select Case mvMaintenanceType
        Case CareServices.XMLMaintenanceControlTypes.xmctAction
          UserHistory.RemoveOtherHistoryNode(HistoryEntityTypes.hetActions, CInt(vList("ActionNumber")))
        Case CareServices.XMLMaintenanceControlTypes.xmctDocument
          UserHistory.RemoveOtherHistoryNode(HistoryEntityTypes.hetDocuments, CInt(vList("DocumentNumber")))
      End Select
      mvRefreshParent = True
      If dgr.Visible = False Then Me.Close() 'Deleted the only item
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      If dgr.Visible Then RePopulateGrid()
    End Try
  End Sub

  Protected Overridable Function ProcessSave(ByVal pDefault As Boolean) As Boolean 'Return true if saved
    Try
      Dim vList As New ParameterList(True)
      If mvContactInfo.ContactNumber > 0 Then vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber

      Dim vEditing As Boolean = EditingExistingRecord()
      If vEditing Then
        'If editing an existing record then get the primary key values
        GetPrimaryKeyValues(vList, mvSelectedRow, True)
      Else
        'For new records add in any additional key values
        GetAdditionalKeyValues(vList)
        If (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocument OrElse mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEMailDocument) AndAlso Not mvDocumentFile Is Nothing Then
          vList("Package") = mvPackage
        End If
        If (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry OrElse mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry) Then
          Dim vGroup As EntityGroup = DataHelper.ContactAndOrganisationGroups(mvContactInfo.ContactGroup)
          If vGroup.AllAddressesUnknown = False And vGroup.UnknownAddress.Length > 0 AndAlso vGroup.UnknownTown.Length > 0 _
          AndAlso epl.GetValue("Address").Length = 0 AndAlso epl.GetValue("Town").Length = 0 AndAlso epl.GetValue("Postcode").Length = 0 Then
            If ShowQuestion(QuestionMessages.qmPostalAddressMissingFormat(vGroup.GroupName), MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
              epl.SetValue("Address", vGroup.UnknownAddress)
              epl.SetValue("Town", vGroup.UnknownTown)
            End If
          End If
        End If
      End If
      If pDefault Then vList("Default") = "Y"
      If epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll) Then
        If mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom Then
          'Need to reverse the Contact and ContactNumber2 values
          Dim vContact As Integer = vList.IntegerValue("ContactNumber")
          vList("ContactNumber") = vList("ContactNumber2")
          vList.IntegerValue("ContactNumber2") = vContact
          '        ElseIf mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers Then
          '          If vList.IntegerValue("AddressNumber") = 0 Then vList("AddressNumber") = ""
        End If
        'Update or Insert record
        If vEditing OrElse mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactNotes OrElse _
                          mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctActionSchedule Then
          mvReturnList = DataHelper.UpdateItem(mvMaintenanceType, vList)
          If vEditing And mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocument Then mvCommsLogSaved = True
        ElseIf mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctMailingOptions Then
          mvReturnList = vList
        Else
          mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
          If Not epl.Recipients Is Nothing AndAlso epl.Recipients.Rows.Count > 1 Then
            For vIndex As Integer = 1 To epl.Recipients.Rows.Count - 1
              vList("ContactNumber") = epl.Recipients.Rows(vIndex).Item("ContactNumber").ToString
              mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
            Next
          End If
          If mvContactDataType = CareServices.XMLContactDataSelectionTypes.xcdtContactInformation Then
            mvContactInfo.ContactNumber = mvReturnList.IntegerValue("ContactNumber")
            mvContactInfo.AddressNumber = mvReturnList.IntegerValue("AddressNumber")
            DoProcessingAfterNewContact()
            DialogResult = Windows.Forms.DialogResult.OK
          Else
            Select Case mvMaintenanceType
              Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
                mvContactInfo.SelectedAddressNumber = mvReturnList.IntegerValue("AddressNumber")
              Case CareServices.XMLMaintenanceControlTypes.xmctLink
                If mvReturnList.ContainsKey("ComplimentaryRelationship") Then
                  If ShowQuestion(QuestionMessages.qmCreateComplimentaryRelationship, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then  'Resource
                    'Need to reverse the Contact and ContactNumber2 values
                    Dim vContact As Integer = vList.IntegerValue("ContactNumber")
                    vList("ContactNumber") = vList("ContactNumber2")
                    vList.IntegerValue("ContactNumber2") = vContact
                    vList("Relationship") = mvReturnList("ComplimentaryRelationship")
                    mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
                  End If
                End If
              Case CareServices.XMLMaintenanceControlTypes.xmctCriterialSet
                mvReturnList("CriteriaSetDesc") = vList("CriteriaSetDesc")        'Add the description to the return list
              Case CareServices.XMLMaintenanceControlTypes.xmctSelectionSet
                UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, mvReturnList.IntegerValue("SelectionSetNumber"), vList("SelectionSetDesc"))
              Case CareServices.XMLMaintenanceControlTypes.xmctDocument, CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
                mvContactInfo.SelectedDocumentNumber = mvReturnList.IntegerValue("DocumentNumber")
                If mvSetRelatedContact IsNot Nothing Then
                  Dim vLinkList As New ParameterList(True)
                  vLinkList.IntegerValue("DocumentNumber") = mvReturnList.IntegerValue("DocumentNumber")
                  vLinkList.IntegerValue("ContactNumber") = mvSetRelatedContact.ContactNumber
                  vLinkList("DocumentLinkType") = "R"
                  DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, vLinkList)
                End If
                If vList.ContainsKey("OurReference") Then mvOurReference = vList("OurReference")
                If Not vEditing Then UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetDocuments, mvContactInfo.SelectedDocumentNumber, mvContactInfo.SelectedDocumentNumber & " - " & mvOurReference)
                CheckForActioners(mvReturnList, Me)
              Case CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
                CheckForActioners(mvReturnList, Me)
              Case CareServices.XMLMaintenanceControlTypes.xmctAction
                mvContactInfo.SelectedActionNumber = mvReturnList.IntegerValue("ActionNumber")
                If mvContactInfo.ContactNumber > 0 Then
                  Dim vLinkList As New ParameterList(True)
                  vLinkList.IntegerValue("ActionNumber") = mvContactInfo.SelectedActionNumber
                  vLinkList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
                  vLinkList("ActionLinkType") = "R"
                  DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vLinkList)
                End If
                If Not vEditing Then UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetActions, mvContactInfo.SelectedActionNumber, vList("ActionDesc").ToString)
            End Select
          End If
          If NoEditingAllowed() Then ProcessNew()
        End If
        mvRefreshParent = True
        If (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocument OrElse mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEMailDocument) AndAlso mvDocumentSaved = False AndAlso Not mvDocumentFile Is Nothing Then
          DataHelper.UpdateDocumentFile(mvContactInfo.SelectedDocumentNumber, mvDocumentFile)
          DataHelper.AddDocumentHistory(CareServices.XMLDocumentHistoryActions.xdhaEdited, mvContactInfo.SelectedDocumentNumber)
          mvDocumentSaved = True
        ElseIf mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContact Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisation Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry Then
          CheckForActioners(mvReturnList, Me)
        ElseIf (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctTCRDocument Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEMailDocument) And Not vEditing Then
          'Add the outcome sub topic
          Dim vOutcome As String = epl.GetValue("Outcome")
          If vOutcome.Length > 0 Then
            Dim vOutComeList As New ParameterList(True)
            vOutComeList("DocumentNumber") = mvReturnList("DocumentNumber")
            If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEMailDocument Then
              vOutComeList("Topic") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.email_in_outcome_topic)
            Else
              If vList("Direction") = "I" Then
                vOutComeList("Topic") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_in_outcome_topic)
              Else
                vOutComeList("Topic") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_outcome_topic)
              End If
            End If
            vOutComeList("SubTopic") = vOutcome
            DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic, vOutComeList)
          End If
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctEMailDocument Then
            'Check if the contact had an existing email address. If not then we must create it now
            If epl.FindPanelControl("Device").Enabled = True Then
              Dim vEMailList As New ParameterList(True)
              vEMailList("ContactNumber") = vList("SenderContactNumber")
              vEMailList("Device") = vList("Device")
              vEMailList("Number") = mvEMailMessage.SenderAddress
              DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctNumber, vEMailList)
            End If
            Dim vLB As CheckedListBox = DirectCast(epl.FindPanelControl("Attachments"), CheckedListBox)
            Dim vDocLinkList1 As New ParameterList(True)
            Dim vDocLinkList2 As New ParameterList(True)
            Dim vNewDocumentNumber As Integer
            For vIndex As Integer = 0 To vLB.Items.Count - 1
              If vLB.GetItemChecked(vIndex) Then
                vList("Package") = EMailApplication.EmailInterface.GetAttachmentPackage(mvEMailMessage, vIndex)
                vList.Remove("DocumentNumber")
                mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
                vNewDocumentNumber = mvReturnList.IntegerValue("DocumentNumber")
                vDocLinkList1.IntegerValue("DocumentNumber") = mvContactInfo.SelectedDocumentNumber
                vDocLinkList1.IntegerValue("DocumentNumber2") = vNewDocumentNumber
                vDocLinkList2.IntegerValue("DocumentNumber") = vNewDocumentNumber
                vDocLinkList2.IntegerValue("DocumentNumber2") = mvContactInfo.SelectedDocumentNumber
                mvReturnList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, vDocLinkList1)
                mvReturnList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, vDocLinkList2)
                DataHelper.UpdateDocumentFile(vNewDocumentNumber, EMailApplication.EmailInterface.GetAttachmentPathName(mvEMailMessage.ID, vIndex))
              End If
            Next
            If mvEMailMessage.DeleteAfterSave Then
              EMailApplication.EmailInterface.ProcessAction(mvEMailMessage, EmailInterface.EMailActions.emaDelete)
              For Each vForm As Form In MDIForm.MdiChildren
                If TypeOf (vForm) Is frmInbox Then
                  DirectCast(vForm, frmInbox).EMailDeleted(mvEMailMessage.ID)
                End If
              Next
            End If
          End If
        End If
        epl.DataChanged = False     'Data saved now
        Return True
      End If
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.imRecordAlreadyExists)
      ElseIf (vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidPositionDates) Then
        ShowInformationMessage(vEx.Message)
      Else
        Throw vEx
      End If
    End Try
  End Function
  Private Function GetApplication(ByVal pNewDocument As Boolean) As ExternalApplication
    Dim vProcessName As String
    Dim vApplication As ExternalApplication

    Dim vComboBox As ComboBox = epl.FindComboBox("Package")
    Dim vRowView As DataRowView = DirectCast(vComboBox.SelectedItem, DataRowView)
    If Not vRowView Is Nothing Then
      mvPackage = vRowView("Package").ToString
      mvExtension = vRowView("DocfileExtension").ToString.ToUpper
      vProcessName = vRowView("ProcessName").ToString.ToLower
      vApplication = FormHelper.GetDocumentApplication(mvExtension)
      If pNewDocument Then AddHandler vApplication.ActionComplete, AddressOf ActionComplete
      epl.EnableControlList("DocumentStyle,StandardDocument", False)
      cmdCreateOrEdit.Enabled = False
      Return vApplication
    Else
      Return Nothing
    End If
  End Function

  Private Sub ActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFilename As String)
    mvDocumentFile = pFilename
    mvDocumentSaved = False
    cmdSave.Enabled = True
    cmdPrint.Enabled = True
    cmdLink1.Enabled = True
    cmdLink2.Enabled = True
  End Sub
  Private Sub ProcessNew()
    dgr.SelectRow(-1)
    mvSelectedRow = -1
    epl.Clear()
    SetDefaults(False)
    SetCommandsForNew()
  End Sub

  Private Sub DoProcessingAfterNewContact()
    'Do activity entry
    FormHelper.ShowDataSheet(Me, frmDataSheet.DataSheetTypes.dstActivities, mvContactInfo, "E", epl.GetValue("Source"), "", "")
    'Do relationship entry
    FormHelper.ShowDataSheet(Me, frmDataSheet.DataSheetTypes.dstRelationships, mvContactInfo, "E", epl.GetValue("Source"), "", "")
    'Do Comms numbers entry if config set
    If mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctContactEntry AndAlso _
       mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry AndAlso _
       AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_enter_numbers) Then
      FormHelper.EditContactData(Me, mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, True)
    End If
    If Me.MdiParent IsNot Nothing Then
      FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, mvContactInfo.ContactNumber, False)
    End If
  End Sub

  Protected Overridable Sub SetCommandsForNew()
    cmdDelete.Enabled = False
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
        If cmdLink2.Enabled Then cmdLink2.Enabled = False 'Disable close site if we are on a new record
      Case CareServices.XMLMaintenanceControlTypes.xmctSuppression
        epl.EnableControl("Suppression", True)
      Case CareServices.XMLMaintenanceControlTypes.xmctLink
        epl.EnableControl("ContactGroup", True)
        epl.EnableControl("ContactNumber2", True)
    End Select
  End Sub

  Private Sub RePopulateGrid(Optional ByVal pNoSelect As Boolean = False)
    'Re-populate grid
    Dim vList As ParameterList = New ParameterList(True)
    GetAdditionalKeyValues(vList)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink
        dgr.Populate(DataHelper.GetActionData(CareServices.XMLActionDataSelectionTypes.xadtActionLinks, mvContactInfo.SelectedActionNumber))
      Case CareServices.XMLMaintenanceControlTypes.xmctActionTopic
        dgr.Populate(DataHelper.GetActionData(CareServices.XMLActionDataSelectionTypes.xadtActionSubjects, mvContactInfo.SelectedActionNumber))
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink
        dgr.Populate(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentLinks, mvContactInfo.SelectedDocumentNumber))
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic
        dgr.Populate(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentSubjects, mvContactInfo.SelectedDocumentNumber))
      Case Else
        dgr.Populate(DataHelper.GetContactData(mvContactDataType, mvContactInfo.ContactNumber, vList))
    End Select
    If dgr.RowCount = 0 Then
      ProcessNew()
    Else
      If dgr.RequiredHeight > splRight.SplitterDistance Then splRight.SplitterDistance = dgr.RequiredHeight
      If pNoSelect = False Then
        'Select current row
        If mvSelectedRow <= 0 Then mvSelectedRow = 0 'TODO Find the records which have just been added
        If mvSelectedRow > dgr.RowCount - 1 Then mvSelectedRow = dgr.RowCount - 1
        dgr.SelectRow(mvSelectedRow)
        SelectRow(mvSelectedRow)
      End If
    End If
  End Sub
  Protected Overridable Sub SetDefaults(Optional ByVal pInitialSetup As Boolean = True)
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctAction
        epl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
      Case CareServices.XMLMaintenanceControlTypes.xmctContact, CareServices.XMLMaintenanceControlTypes.xmctContactEntry
        epl.SetValue("Department", DataHelper.UserInfo.Department)
        If epl.FindTextLookupBox("OwnershipGroup").Text.Length = 0 Then epl.SetValue("OwnershipGroup", DataHelper.UserInfo.OwnershipGroup) 'OwnershipGroup may have been set already by the call to GetBranchFromPostcode
        epl.SetValue("VatCategory", AppValues.ControlValue(AppValues.ControlValues.default_contact_vat_cat))
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry Then
          If AppValues.ControlValue(AppValues.ControlValues.direct_device).Length = 0 Then epl.EnableControl("DirectNumber", False)
          If AppValues.ControlValue(AppValues.ControlValues.switchboard_device).Length = 0 Then epl.EnableControl("SwitchboardNumber", False)
          If AppValues.ControlValue(AppValues.ControlValues.fax_device).Length = 0 Then epl.EnableControl("FaxNumber", False)
          If AppValues.ControlValue(AppValues.ControlValues.mobile_device).Length = 0 Then epl.EnableControl("MobileNumber", False)
          If AppValues.ControlValue(AppValues.ControlValues.email_device).Length = 0 Then epl.EnableControl("EMailAddress", False)
          If AppValues.ControlValue(AppValues.ControlValues.web_device).Length = 0 Then epl.EnableControl("WebAddress", False)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctOrganisation, CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry
        epl.SetValue("Department", DataHelper.UserInfo.Department)
        If epl.FindTextLookupBox("OwnershipGroup").Text.Length = 0 Then epl.SetValue("OwnershipGroup", DataHelper.UserInfo.OwnershipGroup) 'OwnershipGroup may have been set already by the call to GetBranchFromPostcode
        epl.SetValue("Salutation", AppValues.DefaultOrganisationSalutation)
        epl.SetValue("VatCategory", AppValues.ControlValue(AppValues.ControlValues.default_organisation_vat_cat))
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry Then
          If AppValues.ControlValue(AppValues.ControlValues.switchboard_device).Length = 0 Then epl.EnableControl("SwitchboardNumber", False)
          If AppValues.ControlValue(AppValues.ControlValues.fax_device).Length = 0 Then epl.EnableControl("FaxNumber", False)
          If AppValues.ControlValue(AppValues.ControlValues.web_device).Length = 0 Then epl.EnableControl("WebAddress", False)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctMailingOptions
        epl.SetValue("Date", System.DateTime.Now.ToShortDateString)
      Case CareServices.XMLMaintenanceControlTypes.xmctAddresses
        epl.SetValue("Country", AppValues.DefaultCountryCode)
        epl.SetValue("ValidFrom", System.DateTime.Now.ToShortDateString)
      Case CareServices.XMLMaintenanceControlTypes.xmctNumber
        epl.SetValue("AddressNumber", mvContactInfo.AddressNumber.ToString)
      Case CareServices.XMLMaintenanceControlTypes.xmctActivities
        epl.SetValue("ValidFrom", System.DateTime.Now.ToShortDateString)
        epl.SetValue("ValidTo", System.DateTime.Now.AddYears(100).ToShortDateString)
      Case CareServices.XMLMaintenanceControlTypes.xmctDocument, CareServices.XMLMaintenanceControlTypes.xmctTCRDocument, CareServices.XMLMaintenanceControlTypes.xmctEMailDocument
        Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtNewDocumentData, 1))
        epl.SetValue("DocumentNumber", vRow.Item("DocumentNumber").ToString)
        mvOurReference = vRow.Item("OurReference").ToString
        If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctDocument Then
          epl.SetValue("Direction", "O")
          epl.SetValue("Dated", System.DateTime.Now.ToShortDateString)
          epl.SetValue("DocumentStyle", "1", False, True)  'Precis Only
          epl.SetValue("ContactNumber2", DataHelper.UserContactInfo.ContactNumber.ToString)
          epl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
          epl.SetValue("OurReference", vRow.Item("OurReference").ToString)
        ElseIf mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctTCRDocument Then
          epl.SetValue("Direction", "I", False, True)
        Else    'Save EMail
          Dim vContactNo As Integer = DataHelper.FindContactFromEMailAddress(mvEMailMessage.SenderAddress)
          If vContactNo > 0 Then
            mvContactInfo.ContactNumber = vContactNo
            Dim vDevice As String = DataHelper.GetDeviceForEMailAddress(vContactNo, mvEMailMessage.SenderAddress)
            epl.SetValue("Device", vDevice, True)
          End If
          epl.SetValue("Precis", mvEMailMessage.NoteText)
          epl.SetValue("DocumentSubject", TruncateString(mvEMailMessage.Subject, 80))
          If mvEMailMessage.AttachmentCount > 0 Then epl.SetValue("Attachments", mvEMailMessage.AttachmentNameList)
          epl.FindTextLookupBox("Device").SetFilter("EMail = 'Y'")
          mvPackage = DataHelper.GetTextFilePackage
          If mvPackage.Length > 0 Then mvDocumentFile = DataHelper.WriteTempFile(mvEMailMessage.NoteText)
        End If
        If mvContactInfo.ContactNumber > 0 Then
          epl.SetValue("ContactNumber", GetDocumentAddresseeNumber(mvContactInfo).ToString)
        ElseIf mvMaintenanceType <> CareServices.XMLMaintenanceControlTypes.xmctEMailDocument Then
          If mvRelatedContact IsNot Nothing Then
            If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctTCRDocument Then epl.SetValue("Direction", "O")
            epl.SetValue("ContactNumber", mvRelatedContact.ContactNumber.ToString)
          Else
            If MDIForm IsNot Nothing AndAlso MDIForm.CurrentContact IsNot Nothing Then epl.SetValue("ContactNumber", GetDocumentAddresseeNumber(MDIForm.CurrentContact).ToString)
          End If
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctLink
        If Not mvList Is Nothing AndAlso mvList.IntegerValue("ContactNumber2") > 0 Then     'We are restricting to links to this other contact
          Dim vContactInfo As New ContactInfo(mvList.IntegerValue("ContactNumber2"))
          epl.SetValue("ContactGroup", vContactInfo.ContactGroup, True, True)
          epl.SetValue("ContactNumber2", vContactInfo.ContactNumber.ToString, True)
        Else
          epl.SetValue("ContactGroup", mvContactInfo.ContactGroup, False, True)
        End If
        epl.SetValue("ValidFrom", System.DateTime.Now.ToShortDateString)
      Case CareServices.XMLMaintenanceControlTypes.xmctPosition
        epl.SetValue("ValidFrom", System.DateTime.Now.ToShortDateString)
        epl.SetValue("Mail", "Y")
      Case CareServices.XMLMaintenanceControlTypes.xmctRole
        epl.SetValue("ValidFrom", System.DateTime.Now.ToShortDateString)
        If mvContactInfo.ContactType = ContactInfo.ContactTypes.ctContact Then
          epl.SetValue("OrganisationNumber", mvContactInfo.SelectedContactNumber2.ToString, True)
        Else
          epl.SetValue("ContactNumber", mvContactInfo.SelectedContactNumber2.ToString, True)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctSuppression
        epl.SetValue("ValidFrom", System.DateTime.Now.ToShortDateString)
        epl.SetValue("ValidTo", System.DateTime.Now.AddYears(100).ToShortDateString)
      Case CareServices.XMLMaintenanceControlTypes.xmctActionLink
        If pInitialSetup Then
          epl.SetValue("ActionLinkType", "R")
          epl.SetValue("OrganisationContacts", "", True)
          epl.SetValue("PostPoint", "", True)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink
        If pInitialSetup Then
          epl.SetValue("DocumentLinkType", "R")
          epl.SetValue("OrganisationContacts", "", True)
          epl.SetValue("PostPoint", "", True)
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctBatches
        epl.FindTextLookupBox("TransactionType").SetFilter("TransactionSign = 'C' and NegativesAllowed = 'N'")
        Dim vValue As String = epl.FindTextLookupBox("BatchType").GetDataRowItem("DefaultBankAccount")
        If vValue.Length > 0 Then epl.SetValue("BankAccount", vValue, True)
        vValue = GetBatchPaymentMethod(epl.GetValue("BatchType"))
        If vValue.Length > 0 Then epl.SetValue("PaymentMethod", vValue, True)
        If epl.GetValue("BatchType").Length > 0 Then epl.EnableControl("BatchType", False)
        epl.SetValue("BatchDate", System.DateTime.Now.ToShortDateString, Not (AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_set_batch_date_to_tran_date)))
        epl.SetValue("TransactionType", "")
        epl.EnableControl("PayingInSlipNumber", AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_manual_paying_in_slips))
        epl.SetValue("CurrencyCode", AppValues.ControlValue(AppValues.ControlValues.currency_code))
        epl.SetValue("CurrencyExchangeRate", Strings.FormatNumber(1, 3, TriState.False, TriState.False, TriState.False))
        If AppValues.OwnershipMethod = AppValues.OwnershipMethods.omOwnershipGroups Then epl.EnableControl("PrincipalUser", False)
        Select Case epl.GetValue("BatchType")
          Case "CV"
            epl.SetValue("Provisional", "Y", True)
            epl.SetControlVisible("AgencyNumber", True)
            If IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.cv_max_number_of_vouchers)) > 0 Then epl.PanelInfo().PanelItems("NumberOfEntries").MaximumValue = AppValues.ConfigurationValue(AppValues.ConfigurationValues.cv_max_number_of_vouchers)
          Case "CF", "GK", "SR"
            epl.SetValue("Provisional", "Y", True)
        End Select
        epl.EnableControl("Provisional", False)
    End Select
    epl.DataChanged = False
  End Sub

  Private Function GetDocumentAddresseeNumber(ByVal pContactInfo As ContactInfo) As Integer
    If DataHelper.ContactAndOrganisationGroups.ContainsKey(pContactInfo.ContactGroup) Then
      If DataHelper.ContactAndOrganisationGroups(pContactInfo.ContactGroup).AllAddressesUnknown Then
        Dim vRelationship As String = DataHelper.ContactAndOrganisationGroups(pContactInfo.ContactGroup).PrimaryRelationship
        If vRelationship.Length > 0 Then
          Dim vList As New ParameterList(True)
          vList("Relationship") = vRelationship
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, pContactInfo.ContactNumber, vList))
          If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
            mvSetRelatedContact = pContactInfo
            Return CInt(vTable.Rows(0).Item("ContactNumber"))
          End If
        End If
      End If
    End If
    Return pContactInfo.ContactNumber
  End Function

  Private Sub dgr_RowChanging(ByVal pSender As Object, ByRef pCancel As Boolean) Handles dgr.RowChanging
    If epl.DataChanged Then
      pCancel = Not ConfirmCancel()
    End If
  End Sub
  Private Sub dgr_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    'If anything changed then confirm cancel changes
    SelectRow(pRow)
  End Sub
  Private Sub dgr_ContactDropped(ByVal pSender As Object, ByVal pContactInfo As ContactInfo) Handles dgr.ContactDropped
    epl.SetValue("ContactNumber", pContactInfo.ContactNumber.ToString)
    ProcessSave(False)
    RePopulateGrid()
  End Sub

  Private Sub epl_AttachmentNavigate(ByVal pSender As Object, ByVal pIndex As Integer) Handles epl.AttachmentNavigate
    EMailApplication.EmailInterface.ShowAttachment(mvEMailMessage.ID, pIndex)
  End Sub

  Private Sub epl_ContactAddedToOrganisation(ByVal pSender As Object, ByVal pContactNumber As Integer) Handles epl.ContactAddedToOrganisation
    RePopulateGrid(True)      'Repopulate with no selection
    dgr.SelectRow(dgr.FindRow("ContactNumber", pContactNumber.ToString))
  End Sub
  Private Sub epl_ContactDropped(ByVal pSender As Object, ByVal pContactInfo As ContactInfo) Handles epl.ContactDropped
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, CareServices.XMLMaintenanceControlTypes.xmctActionLink
        ProcessSave(False)
        RePopulateGrid()
    End Select
  End Sub
  Private Sub epl_DocumentDropped(ByVal pSender As Object, ByVal pDocumentInfo As DocumentInfo) Handles epl.DocumentDropped
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, CareServices.XMLMaintenanceControlTypes.xmctActionLink
        ProcessSave(False)
        RePopulateGrid()
    End Select
  End Sub
  Private Sub epl_DocumentStyleChanged(ByVal pSender As Object, ByVal pDocumentRequired As Boolean) Handles epl.DocumentStyleChanged
    If EditingExistingRecord() Then
      cmdCreateOrEdit.Enabled = pDocumentRequired
      cmdPrint.Enabled = pDocumentRequired
    Else
      cmdPrint.Enabled = False
      cmdCreateOrEdit.Enabled = pDocumentRequired
      cmdSave.Enabled = Not pDocumentRequired
      cmdLink1.Enabled = Not pDocumentRequired
      cmdLink2.Enabled = Not pDocumentRequired
    End If
  End Sub

  Public WriteOnly Property EMailMessage() As EMailMessage
    Set(ByVal Value As EMailMessage)
      mvEMailMessage = Value
    End Set
  End Property
  Public WriteOnly Property SetModal() As Boolean
    Set(ByVal pvalue As Boolean)
      If pvalue Then
        Me.MdiParent = Nothing
      End If
    End Set
  End Property
  Public WriteOnly Property RelatedContact() As ContactInfo
    Set(ByVal pValue As ContactInfo)
      mvRelatedContact = pValue
    End Set
  End Property
  Public ReadOnly Property ReturnList() As ParameterList
    Get
      Return mvReturnList
    End Get
  End Property
  Public ReadOnly Property DataSelectionType() As CareServices.XMLContactDataSelectionTypes
    Get
      Return mvContactDataType
    End Get
  End Property

  Private Sub epl_EnableCommand(ByVal pSender As Object, ByVal pCommand As EditPanel.MaintenanceCommands, ByVal pEnable As Boolean) Handles epl.EnableCommand
    Select Case pCommand
      Case EditPanel.MaintenanceCommands.mcDefault
        cmdDefault.Enabled = pEnable
    End Select
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As ParameterList) Handles epl.GetCodeRestrictions
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctBatches Then
      If pParameterName = "Product" Then pList("FindProductType") = "B"
    End If
  End Sub

  Private Sub epl_ValidateItem(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String, ByRef pValid As Boolean) Handles epl.ValidateItem
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctBatches
        If pParameterName = "CurrencyCode" Then
          Dim vList As New ParameterList(True)
          vList("CurrencyCode") = pValue
          vList("BatchDate") = epl.GetValue("BatchDate")
          vList("BatchType") = epl.GetValue("BatchType")
          vList.IntegerValue("TraderApplication") = mvList.IntegerValue("TraderApplication")
          Dim vFound As Boolean
          Dim vDataTable As DataTable = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtCurrencyExchangeRates, vList)
          If vDataTable IsNot Nothing Then
            Dim vRow As DataRow = vDataTable.Rows(0)
            If vRow IsNot Nothing Then
              epl.SetValue("CurrencyExchangeRate", vRow.Item("ExchangeRate").ToString)
              If vRow.Item("BankAccount").ToString.Length > 0 Then
                epl.SetValue("BankAccount", vRow.Item("BankAccount").ToString)
              Else
                epl.SetValue("BankAccount", epl.FindTextLookupBox("BatchType").GetDataRowItem("DefaultBankAccount"))
              End If
              vFound = True
            End If
          End If
          If vFound = False Then
            pValid = False
            epl.SetValue("CurrencyExchangeRate", "0.000")
            epl.SetValue("BankAccount", epl.FindTextLookupBox("BatchType").GetDataRowItem("DefaultBankAccount"))
            epl.SetErrorField("CurrencyCode", InformationMessages.imInvalidCurrencyExchangeRate)
          End If
          epl.EnableControl("BankAccount", (epl.GetValue("BankAccount").Length = 0))
        End If
      Case CareServices.XMLMaintenanceControlTypes.xmctContactEntry, CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry
        If pParameterName = "Postcode" Then
          If pValue.Length > 0 Then
            Dim vList As New ParameterList(True)
            vList("Postcode") = pValue
            Dim vFinderType As CareServices.XMLDataFinderTypes
            If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctContactEntry Then
              vFinderType = CareServices.XMLDataFinderTypes.xdftDuplicateContacts
            Else
              vFinderType = CareServices.XMLDataFinderTypes.xdftDuplicateOrganisations
            End If
            Dim vDataSet As DataSet = DataHelper.FindData(vFinderType, vList)
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
            If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
              Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litDuplicateContacts)
              If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                mvContactAddressNumber = vForm.SelectedAddressNumber
                Dim vContactNumber As Integer = CInt(vTable.Rows(vForm.SelectedRow)("ContactNumber"))
                vForm = Nothing
                pValid = False
                If mvContactAddressNumber > 0 Then
                  UseAddressNumber()
                Else
                  epl.DataChanged = False
                  Me.Close()
                  FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, vContactNumber)
                End If
              End If
            End If
          End If
        End If
    End Select
  End Sub

  Private Sub epl_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    Select Case mvMaintenanceType
      Case CareServices.XMLMaintenanceControlTypes.xmctContactEntry, CareServices.XMLMaintenanceControlTypes.xmctOrganisationEntry
        Select Case pParameterName
          Case "Name"
            Dim vList As New ParameterList(True)
            Dim vName As String = DataHelper.GetOrgNameDedupValue(pValue)
            Dim vNotDuplicate As Boolean
            If vName.Length > 0 Then
              vList("Name") = vName & "*"
              vList.IntegerValue("NumberOfRows") = AppValues.MaxDedupRows
              Dim vDataSet As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftDuplicateOrganisations, vList)
              Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
              If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
                Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litDuplicateOrganisations)
                If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                  epl.SetValue("Name", vTable.Rows(vForm.SelectedRow)("Name").ToString)
                  mvOrganisationNumber = CInt(vTable.Rows(vForm.SelectedRow)("OrganisationNumber"))
                  mvContactAddressNumber = vForm.SelectedAddressNumber
                  If mvContactAddressNumber > 0 Then UseAddressNumber()
                  GetOrgNumbers()
                Else
                  mvOrganisationNumber = 0
                  vNotDuplicate = True
                End If
              End If
            Else
              mvOrganisationNumber = 0
              vNotDuplicate = True
            End If
            If vNotDuplicate Then
              mvContactAddressNumber = 0
              epl.ClearControlList("HouseName,Address,Town,County,Postcode,Country")
              epl.EnableControlList("HouseName,Address,Town,County,Postcode,Country", True)
            End If
          Case "Surname"
            Dim vList As New ParameterList(True)
            vList("Surname") = pValue
            vList("ContactGroup") = mvContactInfo.ContactGroup
            Dim vForenames As String = epl.GetValue("Forenames")
            If vForenames.Length > 0 Then vList("Forenames") = FirstWord(vForenames)
            vList.IntegerValue("NumberOfRows") = AppValues.MaxDedupRows
            Dim vDataSet As DataSet = DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftDuplicateContacts, vList)
            Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
            If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
              Dim vForm As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litDuplicateContacts)
              If vForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, CInt(vTable.Rows(vForm.SelectedRow)("ContactNumber")))
                epl.DataChanged = False
                Me.Close()
              End If
            End If
        End Select
      Case CareServices.XMLMaintenanceControlTypes.xmctBatches
        Select Case pParameterName
          Case "OwnershipGroup"
            If pValue.Length = 0 Then epl.SetValue("PrincipalUser", "", True)
            epl.EnableControl("PrincipalUser", (pValue.Length > 0))
        End Select
      Case Else
        If pParameterName = "Direction" Then cmdReply.Enabled = pValue = "I"
    End Select
  End Sub

  Private Sub UseAddressNumber()
    If mvContactAddressNumber > 0 Then
      Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetAddressData(CareServices.XMLAddressDataSelectionTypes.xadtAddressInformation, mvContactAddressNumber))
      epl.SetValue("HouseName", vRow("HouseName").ToString, True)
      epl.SetValue("Address", MultiLine(vRow("Address").ToString), True)
      epl.SetValue("Town", vRow("Town").ToString, True)
      epl.SetValue("County", vRow("County").ToString, True)
      epl.SetValue("Postcode", vRow("Postcode").ToString, True)
      epl.SetValue("Country", vRow("CountryCode").ToString, True)
    Else
      epl.ClearControlList("HouseName,Address,Town,County,Postcode,Country")
      epl.EnableControlList("HouseName,Address,Town,County,Postcode,Country", True)
    End If
  End Sub

  Private Sub GetOrgNumbers()
    epl.EnableControlList("SwitchboardNumber,FaxNumber,WebAddress", True)
    epl.ClearControlList("SwitchboardNumber,FaxNumber,WebAddress")
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, mvOrganisationNumber))
    If vTable IsNot Nothing Then
      Dim vDevice As String
      For Each vRow As DataRow In vTable.Rows
        If IntegerValue(vRow("AddressNumber").ToString) = mvContactAddressNumber OrElse mvContactAddressNumber = 0 Then
          vDevice = vRow("DeviceCode").ToString
          If vDevice = AppValues.ControlValue(AppValues.ControlValues.switchboard_device) Then
            epl.SetValue("SwitchboardNumber", vRow("PhoneNumber").ToString, True)
          ElseIf vDevice = AppValues.ControlValue(AppValues.ControlValues.fax_device) Then
            epl.SetValue("FaxNumber", vRow("PhoneNumber").ToString, True)
          ElseIf vDevice = AppValues.ControlValue(AppValues.ControlValues.web_device) Then
            epl.SetValue("WebAddress", vRow("PhoneNumber").ToString, True)
          End If
        End If
      Next
    End If
  End Sub

  Private Function GetBatchPaymentMethod(ByVal pBatchType As String) As String
    Dim vPayMethod As String = ""

    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctBatches AndAlso pBatchType.Length > 0 Then
      Select Case pBatchType
        Case "CA", "NF"     'Cash, Non-Financial
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash)
        Case "CC"           'CreditCard
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cc)
        Case "CF"           'CAFCards
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_caf_card)
        Case "CS"           'CreditSale
          vPayMethod = AppValues.ControlValue(AppValues.ControlValues.payment_method)
        Case "CV"           'CAFVoucher
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_voucher)
        Case "DC"           'DebitCard
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_dc)
        Case "GK"           'GiftInKind
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_gift_in_kind)
        Case "SO"           'StandingOrder
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_so)
        Case "SP"           'BankStatement
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_sp)
        Case "SR"           'SaleOrReturn
          vPayMethod = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_sr)
      End Select
    End If
    Return vPayMethod
  End Function

End Class


