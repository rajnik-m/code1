Friend Class frmFastDataEntry
  Inherits ThemedForm

  Private Enum EditMode
    Edit
    Run
    Test
  End Enum

  Private mvEditMode As EditMode = EditMode.Run
  Private WithEvents mvPopUpMenu As FastDataEntryMenu
  Private mvContactInfo As ContactInfo
  Private mvDataTableControls As DataTable
  Private mvPageName As String = ""
  Private mvPageNumber As Integer = 0

  'Processing
  Private mvDataChanged As Boolean = False
  Private mvBatchNumber As Integer
  Private mvTransactionNumber As Integer
  Private mvBatchColl As New CollectionList(Of FDEBatchInfo)
  Private mvCMDList As New ParameterList
  Private mvCMDFileName As String
  Private mvPayPlanList As New CollectionList(Of String)
  Private mvCMDPayPlans As New CollectionList(Of String)   'BR13524: Only used for creating CMD for multiple PP
  Private mvMultiplePages As Boolean
  Private mvFormClosingAllowed As Boolean = True
  Private mvTotalDonationAmount As Double
  Private mvPreviousCardNumberNotRequired As Boolean = False

#Region " Initialisation "

  Public Sub New(ByVal pPageNumber As Integer, ByVal pTesting As Boolean)
    'This will be used when running the FastDataEntry page from the Applications menu
    'This is also used when selecting to test the FastDataEntry page from Data Entry Application Maintenance

    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    MainHelper.SetMDIParent(Me)
    If pTesting Then mvEditMode = EditMode.Test
    Dim vList As New ParameterList(True, True)
    vList.IntegerValue("FdePageNumber") = pPageNumber
    Dim vDT As DataTable = DataHelper.GetFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePages, vList)
    If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
      Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vDT.DataSet)
      mvPageNumber = IntegerValue(vRow.Item("FdePageNumber").ToString)
      mvPageName = vRow.Item("FdePageName").ToString
      Me.Size = New Size(IntegerValue(vRow("FdePageWidth").ToString), IntegerValue(vRow("FdePageHeight").ToString))
      If Me.Parent IsNot Nothing AndAlso Me.Size.Width > Me.Parent.ClientSize.Width Then Me.Size = New Size(Me.Parent.ClientSize.Width, Me.Height)
      Me.Text = vRow.Item("fdePageTitle").ToString
      InitialiseControls(True)
    End If
    If Me.MdiParent IsNot Nothing Then Me.MdiParent = Nothing
  End Sub

  Public Sub New(ByVal pRow As DataRow, ByVal pEditing As Boolean)
    'This will be used when editing an existing FastDataEntry page

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    If pEditing Then mvEditMode = EditMode.Edit
    mvPageNumber = IntegerValue(pRow.Item("FdePageNumber").ToString)
    mvPageName = pRow.Item("FdePageName").ToString
    Me.Size = New Size(IntegerValue(pRow("FdePageWidth").ToString), IntegerValue(pRow("FdePageHeight").ToString))
    Me.Text = pRow.Item("fdePageTitle").ToString
    InitialiseControls(True)
  End Sub

  Public Sub New(ByVal pPageTitle As String, ByVal pPageName As String)
    'This will be used when creating a new FdePage

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    mvEditMode = EditMode.Edit
    mvPageName = pPageName
    mvPageNumber = 0
    Me.Text = pPageTitle
    InitialiseControls(False)
    'Fist thing to do is save the page
    SaveFDEPage()
  End Sub

  Private Sub InitialiseControls(ByVal pGetControls As Boolean)
    SetControlTheme()
    cmdNext.Visible = (mvEditMode = EditMode.Run)
    cmdOK.Visible = (mvEditMode = EditMode.Run)
    cmdTest.Visible = (mvEditMode = EditMode.Edit)
    If mvEditMode <> EditMode.Run Then cmdCancel.Text = ControlText.CmdClose

    If mvEditMode <> EditMode.Run Then CheckFDEControls()

    If pGetControls Then
      'Populate the Form with the existing controls
      Dim vList As New ParameterList(True, True)
      vList.IntegerValue("FdePageNumber") = mvPageNumber
      Dim vAddTransDetails As CareFDEControl = Nothing
      Dim vTelemarketing As CareFDEControl = Nothing
      Dim vDT As DataTable = DataHelper.GetFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePageControls, vList)
      If vDT IsNot Nothing Then
        For Each vRow As DataRow In vDT.Rows
          Dim vItem As CareFDEControl = Nothing
          Select Case vRow("FdeUserControl").ToString
            Case "ACTIVITYDISPLAY"
              vItem = New FDEActivityDisplay(CareNetServices.FDEControlTypes.ActivityDisplay, vRow, (mvEditMode = EditMode.Edit))
            Case "ADDDONATIONCC"
              vItem = New FDEAddDonationCC(CareNetServices.FDEControlTypes.AddDonationCC, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.SelectedContactChanged, AddressOf CareFDEControl_SelectedContactChanged
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged
              AddHandler vItem.ReferenceMandatory, AddressOf CareFDEControl_ReferenceMandatory
            Case "PRODUCTSALE"
              vItem = New FDEAddDonationCC(CareNetServices.FDEControlTypes.ProductSale, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.SelectedContactChanged, AddressOf CareFDEControl_SelectedContactChanged
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged
              AddHandler vItem.ReferenceMandatory, AddressOf CareFDEControl_ReferenceMandatory
            Case "ADDMEMBERDD"
              vItem = New FDEAddMemberDD(CareNetServices.FDEControlTypes.AddMemberDD, vRow, (mvEditMode = EditMode.Edit))
            Case "ADDREGULARDON"
              vItem = New FDEAddRegularDonation(CareNetServices.FDEControlTypes.AddRegularDonation, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.RegDonationAdded, AddressOf CareFDEControl_RegDonationAdded
            Case "ADDRESSDISPLAY"
              vItem = New FDEAddressDisplay(CareNetServices.FDEControlTypes.AddressDisplay, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.AddressChanged, AddressOf CareFDEControl_AddressChanged
            Case "ADDTRANSACTIONDETAILS"
              vItem = New FDEAddTransactionDetails(CareNetServices.FDEControlTypes.AddTransactionDetails, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.TransactionDateChanged, AddressOf careFDEControl_TransactionDateChanged
              AddHandler vItem.SourceChanged, AddressOf CareFDEControl_SourceChanged
              If mvEditMode <> EditMode.Edit Then vAddTransDetails = vItem
            Case "COMMUNICATIONSDISPLAY"
              vItem = New FDECommunicationsDisplay(CareNetServices.FDEControlTypes.CommunicationsDisplay, vRow, (mvEditMode = EditMode.Edit))
            Case "CONTACTSELECTION"
              vItem = New FDEContactSelection(CareNetServices.FDEControlTypes.ContactSelection, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged
              AddHandler vItem.BankDetailsChange, AddressOf CareFDEControl_ClearBankDetails
            Case "DISPLAYLABEL"
              vItem = New FDEDisplayLabel(CareNetServices.FDEControlTypes.DisplayLabel, vRow, (mvEditMode = EditMode.Edit))
            Case "GIFTAIDDISPLAY"
              vItem = New FDEGiftAidDisplay(CareNetServices.FDEControlTypes.GiftAidDisplay, vRow, (mvEditMode = EditMode.Edit))
            Case "SUPPRESSIONDISPLAY"
              vItem = New FDESuppressionDisplay(CareNetServices.FDEControlTypes.SuppressionDisplay, vRow, (mvEditMode = EditMode.Edit))
            Case "TELEMARKETING"
              vItem = New FDETelemarketing(CareNetServices.FDEControlTypes.Telemarketing, vRow, (mvEditMode = EditMode.Edit))
              AddHandler vItem.SourceChanged, AddressOf CareFDEControl_SetAndRefreshSource  'to update source in AddTransactionDetails module
              AddHandler vItem.SelectedContactChanged, AddressOf CareFDEControl_SelectedContactChanged  'to update contact number in ContactSelection module
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged  'to update contact number in all other modules
              AddHandler vItem.EnableOtherModules, AddressOf CareFDEControl_EnableOtherModules  'to disable all other modules
              AddHandler vItem.FormClosingAllowed, AddressOf CareFDEControl_FormClosingAllowed  'to not allow closing the form when a call has been made
              vTelemarketing = vItem
          End Select
          If vItem IsNot Nothing Then AddFDEControl(vItem)
        Next
      End If
      If vAddTransDetails IsNot Nothing Then vAddTransDetails.SetDefaults() 'To raise SourceChanged event to set Source details on AddDonationCC and AddMemberDD controls
      If vTelemarketing IsNot Nothing Then
        'To raise SourceChanged event to set Source details on AddTransactionDetails module which will then reset the source on other required modules
        Dim vSource As String = vTelemarketing.epl.FindTextLookupBox("Segment").GetDataRowItem("Source")
        If vSource.Length > 0 Then CareFDEControl_SetAndRefreshSource(vTelemarketing, vSource, "", "")
        CareFDEControl_EnableOtherModules(vTelemarketing, False) 'force this event as it has not been executed from SetDefaults
      End If
    End If
    Select Case mvEditMode
      Case EditMode.Edit
        mvPopUpMenu = New FastDataEntryMenu()
        Me.ContextMenuStrip = mvPopUpMenu
      Case EditMode.Run
        'Reset DataChanged to False
        For Each vControl As Control In pnl.Controls
          If TypeOf (vControl) Is CareFDEControl Then
            CType(vControl, CareFDEControl).ResetDataChanged()
          End If
        Next
        mvDataChanged = False
    End Select
  End Sub

  Private Sub CheckFDEControls()
    Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtFastDataEntryUserControls)
    Dim vUpdate As Boolean

    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
      vUpdate = True
    Else
      'Check for a specific Control missing
      Dim vColumns() As DataColumn = {vDT.Columns("FDEUserControl")}
      vDT.PrimaryKey = vColumns
      Dim vRow As DataRow = vDT.Rows.Find("TELEMARKETING")
      If vRow Is Nothing Then
        vUpdate = True
      Else
        vUpdate = Not vRow("DefaultParameters").ToString.Contains("CardNumberNotRequired")
      End If
    End If

    If vUpdate Then
      vDT = New DataTable("FDEUserControls")
      Dim vValues As String(,) = { _
      {"CONTACTSELECTION", "Contact Selection", "FCSE", "", ""}, _
      {"ADDRESSDISPLAY", "Address Display", "FASE", "", ""}, _
      {"SUPPRESSIONDISPLAY", "Suppression Display", "FSDI", "DisplayControlType", "SuppressionGroup"}, _
      {"ACTIVITYDISPLAY", "Activity Display", "FADI", "DisplayControlType", "ActivityGroup"}, _
      {"COMMUNICATIONSDISPLAY", "Communications Display", "FCDI", "", ""}, _
      {"GIFTAIDDISPLAY", "Gift Aid Display", "FGDI", "", ""}, _
      {"DISPLAYLABEL", "Display Label", "FLBL", "DisplayText", "FontName,FontSize,FontStyle"}, _
      {"ADDTRANSACTIONDETAILS", "Add Transaction Details", "FTRD", "", "Source,TransactionOrigin"}, _
      {"ADDDONATIONCC", "Add Donation", "FDCC", "OnlineCCAuthorisation,BankAccount", "Product,Rate,DistributionCodeLookupGroup,PaymentMethod,CardNumberNotRequired"}, _
      {"ADDMEMBERDD", "Add Member", "FMDD", "AlwaysAddDDTransaction", "MembershipType,PaymentFrequency,DistributionCodeLookupGroup,BankAccount,ClaimDay,MandateType,Branch"}, _
      {"ADDREGULARDON", "Add Regular Donations", "FRDD", "", "Product,Rate,PaymentFrequency,BankAccount,ClaimDay,MandateType"}, _
      {"TELEMARKETING", "Telemarketing", "FRTM", "", "Campaign,Appeal,Segment,OverwriteCallBackTime"} _
      }
      Dim vFields As String() = {"FdeUserControl", "ControlTitle", "FpPageType", "InitialParameters", "DefaultParameters"}
      Dim vParams As New ParameterList(True)

      For vIndex As Integer = 0 To vValues.GetLength(0) - 1
        For vItem As Integer = 0 To vValues.GetLength(1) - 1
          'Debug.Print(vValues(vIndex, vItem))
          vParams(vFields(vItem)) = vValues(vIndex, vItem)
        Next
        DataHelper.UpdateFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdeUserControls, vParams)
      Next
    End If
  End Sub

#End Region

#Region " Public Properties "
  Friend ReadOnly Property PageNumber() As Integer
    Get
      Return mvPageNumber
    End Get
  End Property
#End Region

#Region " Menu Handling "

  Private Sub mvPopUpMenu_MenuSelected(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FastDataEntryMenu.FDEBrowserMenuItems) Handles mvPopUpMenu.MenuSelected
    If pItem = FastDataEntryMenu.FDEBrowserMenuItems.EditPage Then
      'Edit Page
      Dim vDefaults As New ParameterList
      vDefaults("FdePageName") = mvPageName
      vDefaults("FdePageTitle") = Me.Text
      Dim vFrmAP As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptAddFastDataEntryPage, vDefaults, Nothing)
      If vFrmAP.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
        Dim vReturn As ParameterList = vFrmAP.ReturnList
        mvPageName = vReturn("FdePageName")
        Me.Text = vReturn("FdePageTitle")
        mvDataChanged = True
      End If
    ElseIf pItem >= FastDataEntryMenu.FDEBrowserMenuItems.ActivityDisplay Then
      'Add Module
      If mvDataTableControls Is Nothing Then mvDataTableControls = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtFastDataEntryUserControls)
      Dim vUserControl As String = pMenuItem.Name.Substring(3)
      Dim vRow As DataRow = Nothing
      If mvDataTableControls IsNot Nothing Then
        For Each vRow In mvDataTableControls.Rows
          If vRow.Item("FdeUserControl").ToString = vUserControl Then Exit For
        Next
      End If
      If vRow IsNot Nothing Then
        'Dim vPageType As String = vRow("FpPageType").ToString
        Dim vInitialParameters As String = vRow("InitialParameters").ToString
        Dim vDefaultParameters As String = vRow("DefaultParameters").ToString
        Dim vInitialSettings As String = ""
        Dim vDefaultSettings As String = ""
        Dim vAddControl As Boolean = True
        If vInitialParameters.Length + vDefaultParameters.Length > 0 Then
          vAddControl = DataHelper.DisplayControlParameters(Me, vUserControl, vInitialParameters, vDefaultParameters, vInitialSettings, vDefaultSettings)
        End If
        If vAddControl Then
          Dim vItem As CareFDEControl = Nothing
          Dim vSequenceNumber As Integer = pnl.Controls.Count + 1
          Select Case pItem
            Case FastDataEntryMenu.FDEBrowserMenuItems.ActivityDisplay
              vItem = New FDEActivityDisplay(CareNetServices.FDEControlTypes.ActivityDisplay, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.AddDonationCC
              vItem = New FDEAddDonationCC(CareNetServices.FDEControlTypes.AddDonationCC, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
              AddHandler vItem.SelectedContactChanged, AddressOf careFDEControl_SelectedContactChanged
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged
            Case FastDataEntryMenu.FDEBrowserMenuItems.AddMemberDD
              vItem = New FDEAddMemberDD(CareNetServices.FDEControlTypes.AddMemberDD, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.AddRegularDonation
              vItem = New FDEAddRegularDonation(CareNetServices.FDEControlTypes.AddRegularDonation, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
              AddHandler vItem.RegDonationAdded, AddressOf CareFDEControl_RegDonationAdded
            Case FastDataEntryMenu.FDEBrowserMenuItems.AddressDisplay
              vItem = New FDEAddressDisplay(CareNetServices.FDEControlTypes.AddressDisplay, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
              AddHandler vItem.AddressChanged, AddressOf CareFDEControl_AddressChanged
            Case FastDataEntryMenu.FDEBrowserMenuItems.AddTransaction
              vItem = New FDEAddTransactionDetails(CareNetServices.FDEControlTypes.AddTransactionDetails, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.CommunicationsDisplay
              vItem = New FDECommunicationsDisplay(CareNetServices.FDEControlTypes.CommunicationsDisplay, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.ContactSelection
              vItem = New FDEContactSelection(CareNetServices.FDEControlTypes.ContactSelection, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged
              AddHandler vItem.BankDetailsChange, AddressOf CareFDEControl_ClearBankDetails
            Case FastDataEntryMenu.FDEBrowserMenuItems.DisplayLabel
              vItem = New FDEDisplayLabel(CareNetServices.FDEControlTypes.DisplayLabel, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.GiftAidDisplay
              vItem = New FDEGiftAidDisplay(CareNetServices.FDEControlTypes.GiftAidDisplay, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.SuppressionDisplay
              vItem = New FDESuppressionDisplay(CareNetServices.FDEControlTypes.SuppressionDisplay, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.Telemarketing
              vItem = New FDETelemarketing(CareNetServices.FDEControlTypes.Telemarketing, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
            Case FastDataEntryMenu.FDEBrowserMenuItems.ProductSale
              vItem = New FDEAddDonationCC(CareNetServices.FDEControlTypes.ProductSale, vRow, vInitialSettings, vDefaultSettings, mvPageNumber, vSequenceNumber, True)
              AddHandler vItem.SelectedContactChanged, AddressOf CareFDEControl_SelectedContactChanged
              AddHandler vItem.ContactChanged, AddressOf CareFDEControl_ContactChanged
          End Select
          If vItem IsNot Nothing Then
            AddFDEControl(vItem)
          End If
        End If
      End If
    End If
  End Sub

#End Region

#Region " Private Methods "

  Private Sub frmFastDataEntry_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    MainHelper.EnableTraderApplications(True)
  End Sub

  Private Sub frmFastDataEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    If mvFormClosingAllowed Then
      If mvDataChanged = False Then mvDataChanged = ControlDataChanged()
      If mvDataChanged Then
        Dim vCancel As Boolean = Not (ConfirmCancel())
        If vCancel Then
          e.Cancel = True
        Else
          For Each vControl As Control In pnl.Controls
            If TypeOf (vControl) Is FDETelemarketing Then
              Dim vFDEControl As FDETelemarketing = DirectCast(vControl, FDETelemarketing)
              If vFDEControl.ProcessCancel Then
                DataHelper.UpdateTelemarketingContact(CareNetServices.TelemarketingUpdateType.Cancel, IntegerValue(vFDEControl.epl.FindTextLookupBox("Segment").GetDataRowItem("SelectionSet")), mvContactInfo.ContactNumber)
              End If
            End If
          Next
        End If
      End If
    Else
      e.Cancel = True
      ShowInformationMessage(InformationMessages.ImMustSubmitData)
    End If
  End Sub

  Private Sub frmFastDataEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    bpl.RepositionButtons()
    ResizeControls()
  End Sub

  Private Sub frmFastDataEntry_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    If mvEditMode = EditMode.Edit Then ResizeControls()
  End Sub

  Private Sub ResizeControls()
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        DirectCast(vControl, CareFDEControl).ResizeControl(Me.Width)
      End If
    Next
  End Sub

  Private Sub AddFDEControl(ByVal pItem As CareFDEControl)
    pnl.Visible = False
    AddHandler pItem.DeleteControl, AddressOf CareFDEControl_DeleteControl
    pnl.Controls.Add(pItem)
    pItem.Dock = DockStyle.Top
    'pItem.Anchor = CType(AnchorStyles.Left + AnchorStyles.Right, AnchorStyles)
    pItem.BringToFront()
    pItem.ResizeControl(Me.Width)
    'mvDataChanged = True
    pnl.Visible = True
  End Sub

  Private Function ControlDataChanged() As Boolean
    Dim vChanged As Boolean
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        If DirectCast(vControl, CareFDEControl).ControlSizeChanged Then vChanged = True
      End If
      If vChanged Then Exit For
    Next
    If vChanged = False AndAlso mvEditMode = EditMode.Run Then
      Dim vGotFinancialControls As Boolean = False
      For Each vControl As Control In pnl.Controls
        If TypeOf (vControl) Is CareFDEControl Then
          Select Case CType(vControl, CareFDEControl).ControlType
            Case CareNetServices.FDEControlTypes.ActivityDisplay, CareNetServices.FDEControlTypes.AddressDisplay, CareNetServices.FDEControlTypes.CommunicationsDisplay, _
                 CareNetServices.FDEControlTypes.GiftAidDisplay, CareNetServices.FDEControlTypes.SuppressionDisplay
              'Any changes to these controls are applied immediately so ignore DataChanged
            Case CareNetServices.FDEControlTypes.ContactSelection
              'If a contact has been selected ignore DataChanged unless there are financial controls
              If vGotFinancialControls = True AndAlso DirectCast(vControl, CareFDEControl).epl.DataChanged = True Then vChanged = True
            Case Else
              vGotFinancialControls = True
              If DirectCast(vControl, CareFDEControl).epl.DataChanged Then vChanged = True
          End Select
        End If
        If vChanged Then Exit For
      Next
    End If
    Return vChanged
  End Function

  Private Sub SaveFDEPage()
    Dim vList As New ParameterList(True, True)
    If mvPageNumber > 0 Then vList.IntegerValue("FdePageNumber") = mvPageNumber
    vList("FdePageTitle") = Me.Text
    vList("FdePageName") = mvPageName
    vList.IntegerValue("FdePageHeight") = Me.Height
    vList.IntegerValue("FdePageWidth") = Me.Width
    Dim vResultList As ParameterList = DataHelper.UpdateFastDataEntryData(CareNetServices.XMLFastDataEntryTypes.fdePages, vList)
    mvPageNumber = vResultList.IntegerValue("FdePageNumber")
  End Sub

  Private Sub SaveFDEPageItems()
    Dim vSequenceNumber As Integer = 0
    For vIndex As Integer = (pnl.Controls.Count - 1) To 0 Step -1
      'Controls are in the collection in reverse order
      Dim vControl As Control = pnl.Controls(vIndex)
      If TypeOf (vControl) Is CareFDEControl Then
        vSequenceNumber += 1
        DirectCast(vControl, CareFDEControl).Save(vSequenceNumber)
      End If
    Next
  End Sub

  Private Function ProcessSubmit(ByVal pNext As Boolean) As Boolean
    'Called from Next & Submit buttons
    Dim vValid As Boolean
    Dim vList As New ParameterList(True, True)
    Dim vModulesSubmitted As Boolean = False
    Dim vCardNumberNotRequired As Boolean
    Dim vIsCardTransaction As Boolean = False
    Dim vAddDonation As Boolean
    If mvMultiplePages = False AndAlso pNext = True Then mvMultiplePages = True

    'Validate (Build parameter list from all relevant modules)
    Dim vGotContactModule As Boolean = False
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        Select Case vFDEControl.ControlType
          Case CareNetServices.FDEControlTypes.ContactSelection
            vValid = vFDEControl.BuildParameterList(vList)
            vGotContactModule = True
        End Select
        If TypeOf vFDEControl Is FDEAddDonationCC Then
          vIsCardTransaction = vFDEControl.IsCardTransaction
        End If
      End If
    Next
    If vGotContactModule = False Then Throw New CareException(CareException.ErrorNumbers.enNoContactSelected)

    Dim vAddPaymentPlan As Boolean
    If vValid Then
      For Each vControl As Control In pnl.Controls
        If TypeOf (vControl) Is CareFDEControl Then
          Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
          Select Case vFDEControl.ControlType
            Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.AddMemberDD, CareNetServices.FDEControlTypes.AddTransactionDetails, _
                 CareNetServices.FDEControlTypes.AddRegularDonation, CareNetServices.FDEControlTypes.Telemarketing, CareNetServices.FDEControlTypes.ProductSale
              vList = New ParameterList(True, True)
              vValid = vFDEControl.BuildParameterList(vList)
              If vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddDonationCC Then
                vAddDonation = True
              ElseIf vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddMemberDD OrElse vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddRegularDonation Then
                vAddPaymentPlan = True
              End If
          End Select
        End If
        If vValid = False Then Exit For
      Next
    End If

    'See if we need to process any incentives
    If vValid Then
      Dim vCheckIncentives As Boolean
      Dim vIncList As New ParameterList()
      For vIndex As Integer = pnl.Controls.Count - 1 To 0 Step -1
        vCheckIncentives = False
        If TypeOf (pnl.Controls(vIndex)) Is CareFDEControl Then
          Dim vFdeControl As CareFDEControl = DirectCast(pnl.Controls(vIndex), CareFDEControl)
          vIncList = New ParameterList()
          Select Case vFdeControl.ControlType
            Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.AddMemberDD, CareNetServices.FDEControlTypes.AddRegularDonation, CareNetServices.FDEControlTypes.ProductSale
              vCheckIncentives = vFdeControl.CheckIncentives(vIncList)
              If vCheckIncentives Then
                vIncList.AddConnectionData()
                Dim vForm As New frmIncentives
                Dim vDS As DataSet = vForm.GetIncentivesData(vIncList, True, True, True, False)
                Dim vDT As DataTable = Nothing
                If vDS IsNot Nothing Then vDT = DataHelper.GetTableFromDataSet(vDS)
                If vDT IsNot Nothing Then
                  Dim vSequence As New StringBuilder
                  Dim vQuantity As New StringBuilder
                  For Each vRow As DataRow In vDT.Rows
                    If vSequence.Length > 0 Then vSequence.Append(",")
                    If vQuantity.Length > 0 Then vQuantity.Append(",")
                    vSequence.Append(vRow("SequenceNumber").ToString)
                    vQuantity.Append(vRow("Quantity").ToString)
                  Next
                  If vSequence.Length > 0 AndAlso vQuantity.Length > 0 Then
                    'Add the SequenceNumber & Quantity lists to the Control so they get submitted with the data
                    vFdeControl.AddIncentives(vSequence.ToString, vQuantity.ToString)
                  End If
                End If
              End If
          End Select
        End If
      Next
    End If

    'Now submit the modules
    Dim vTransList As New ParameterList
    Dim vTelemarketing As Boolean = False
    If vValid Then
      For Each vControl As Control In pnl.Controls
        If TypeOf (vControl) Is CareFDEControl Then
          Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
          If vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddTransactionDetails Then
            vFDEControl.BuildParameterList(vTransList)
          End If
        End If
      Next
      If vTransList.Count = 0 Then
        'See if there are any controls that need submitting
        For Each vControl As Control In pnl.Controls
          If TypeOf (vControl) Is CareFDEControl Then
            Select Case DirectCast(vControl, CareFDEControl).ControlType
              Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.AddMemberDD, CareNetServices.FDEControlTypes.AddRegularDonation, CareNetServices.FDEControlTypes.ProductSale
                vValid = False
            End Select
          End If
        Next
        If vValid = False Then Throw New CareException(CareException.ErrorNumbers.enNoTransactionDetails)
      End If

      Dim vReturnDS As DataSet = Nothing
      For vIndex As Integer = pnl.Controls.Count - 1 To 0 Step -1
        If TypeOf (pnl.Controls(vIndex)) Is CareFDEControl Then
          Dim vCashPP As Boolean = False
          Dim vBatchType As String = "CA"
          Dim vBatchInfo As New FDEBatchInfo
          Dim vFDEControl As CareFDEControl = DirectCast(pnl.Controls(vIndex), CareFDEControl)
          vList = New ParameterList(True, True)
          Select Case vFDEControl.ControlType
            Case CareNetServices.FDEControlTypes.AddDonationCC, CareNetServices.FDEControlTypes.ProductSale
              If vFDEControl.CanSubmit() Then
                vList.FillFromValueList(vTransList.ValueList)
                vFDEControl.BuildParameterList(vList)
                vList.Add("AccountName", mvContactInfo.ContactName)
                Dim vCCNumber As String = ""
                If vList.ContainsKey("CreditCardNumber") Then
                  vCCNumber = vList("CreditCardNumber")
                  vList.Remove("CreditCardNumber")
                End If
                If mvBatchColl.Count > 0 Then
                  For Each vBI As FDEBatchInfo In mvBatchColl
                    If vBI.BatchType = vBatchType Then
                      vList.IntegerValue("BatchNumber") = vBI.BatchNumber
                      vList.IntegerValue("TransactionNumber") = vBI.TransactionNumber
                      Exit For
                    End If
                  Next
                End If
                If vList.Contains("CardNumberNotRequired") AndAlso vList("CardNumberNotRequired") = "Y" Then
                  vCardNumberNotRequired = True
                  mvPreviousCardNumberNotRequired = True
                End If
                If vFDEControl.HasCompleted = False Then
                  If vList("PaymentMethod").ToString = AppValues.ConfigurationValue(AppValues.ConfigurationValues.pm_cash, "CASH") Then vValid = IsTotalDonationValid(mvTotalDonationAmount + DoubleValue(vList("Amount").ToString), mvContactInfo)
                  If vValid Then
                    mvTotalDonationAmount += DoubleValue(vList("Amount").ToString)
                    vReturnDS = DataHelper.AddProductSale(vList)
                    vFDEControl.ResetIncentives()
                    vFDEControl.HasCompleted = True
                    If vCCNumber.Length > 0 Then vList("CreditCardNumber") = vCCNumber
                  End If
                End If
              ElseIf vIsCardTransaction AndAlso vCardNumberNotRequired = False AndAlso mvPreviousCardNumberNotRequired Then
                'BR21572: Where module is no longer possible to submit but previously was
                'and the 'Card Number Not Required' checkbox was set, set vCardNumberNotRequired to True such that this module will be processed without card number.
                vCardNumberNotRequired = True
              End If

            Case CareNetServices.FDEControlTypes.AddMemberDD, CareNetServices.FDEControlTypes.AddRegularDonation
              If vFDEControl.CanSubmit() Then
                vList.FillFromValueList(vTransList.ValueList)
                If vList.Contains("Reference") Then vList.Remove("Reference")
                vFDEControl.BuildParameterList(vList)
                vBatchType = "NF"
                If vList.ContainsKey("BatchNumber") Then
                  vList.Remove("BatchNumber")
                  vList.Remove("TransactionNumber")
                End If
                If vList.ContainsKey("AccountNumber") = False Then vCashPP = True
                If mvBatchColl.Count > 0 Then
                  For Each vBI As FDEBatchInfo In mvBatchColl
                    If vBI.BatchType = "NF" AndAlso vBI.BatchNumber > 0 Then
                      vList.IntegerValue("BatchNumber") = vBI.BatchNumber
                      vList.IntegerValue("TransactionNumber") = vBI.TransactionNumber
                    End If
                  Next
                End If
                If vFDEControl.HasCompleted = False Then
                  Try
                    If vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddMemberDD Then
                      vReturnDS = DataHelper.AddMembership(vList)
                    Else
                      If vCashPP Then
                        vReturnDS = DataHelper.AddPaymentPlan(CareNetServices.ppType.pptOther, vList)
                      Else
                        vReturnDS = DataHelper.AddPaymentPlan(CareNetServices.ppType.pptDD, vList)
                      End If
                    End If
                    If pNext Then vFDEControl.ResetIncentives()
                    vFDEControl.HasCompleted = True
                  Catch vCareEx As CareException
                    If vCareEx.ErrorNumber = CareException.ErrorNumbers.enDDReferenceSixAlphas Then
                      ShowErrorMessage(vCareEx.Message)
                      Return False
                    Else
                      Throw vCareEx
                    End If
                  Catch vEx As Exception
                    Throw vEx
                  End Try
                End If
              End If
            Case CareNetServices.FDEControlTypes.Telemarketing
              vTelemarketing = True
              If vFDEControl.HasCompleted = False AndAlso pNext = False Then
                vFDEControl.BuildParameterList(vList)
                vList.IntegerValue("AddresseeContactNumber") = mvContactInfo.ContactNumber
                vList.IntegerValue("AddresseeAddressNumber") = mvContactInfo.AddressNumber
                vList.IntegerValue("SenderContactNumber") = DataHelper.UserContactInfo.ContactNumber
                vList.IntegerValue("SenderAddressNumber") = DataHelper.UserContactInfo.AddressNumber
                Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtNewDocumentData, 1))
                Dim vDocumentNumber As String = vRow.Item("DocumentNumber").ToString
                vList("DocumentNumber") = vDocumentNumber
                vList("DocumentClass") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_document_class)
                vList("DocumentType") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_document_type)
                vList("Dated") = AppValues.TodaysDate
                vList("CallDuration") = vFDEControl.epl.GetValue("CallDuration")
                vList("TotalDuration") = vFDEControl.epl.GetValue("TotalDuration")
                vList("Direction") = "O"
                vList("Source") = vFDEControl.epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("Source")
                vList("SelectionSet") = vFDEControl.epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("SelectionSet")

                'override topic and sub topic as by default it will be read from the datasheet
                vRow = vFDEControl.epl.FindPanelControl(Of TextLookupBox)("Campaign").GetDataRow
                Dim vOutcomeTopic As String = ""
                If vRow.Table.Columns.Contains("Topic") Then vOutcomeTopic = vRow("Topic").ToString
                If vOutcomeTopic.Length = 0 Then vOutcomeTopic = AppValues.ConfigurationValue(AppValues.ConfigurationValues.phone_out_outcome_topic)
                vList("Topic") = vOutcomeTopic
                Dim vOutcome As String = vList("Outcome")
                vList("SubTopic") = vOutcome

                'Remove unwanted params
                vList.Remove("Outcome")
                vList.Remove("TopicGroup")
                vList.Remove("TopicDataSheet")
                vList.Remove("Camapign")
                vList.Remove("Appeal")
                vList.Remove("Segment")
                Dim vCallBackTime As String = vList("CallBackTime").ToString
                vList.Remove("CallBackTime")
                vList.Remove("OverwriteCallBackTime")

                'save the communication log record
                DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctTCRDocument, vList)

                'save any topic and sub topic records
                Dim vDataSheet As TopicDataSheet = DirectCast(FindControl(vFDEControl.epl, "TopicDataSheet", False), TopicDataSheet)
                vDataSheet.SaveTopics(IntegerValue(vDocumentNumber), vList("Topic").ToString, vList("SubTopic").ToString)

                vList = New ParameterList(True)
                vList("SelectionSet") = vFDEControl.epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("SelectionSet")
                vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
                vList("Topic") = vOutcomeTopic
                vList("SubTopic") = vOutcome
                vList("CallBackTime") = vCallBackTime
                DataHelper.UpdateTelemarketingContact(CareNetServices.TelemarketingUpdateType.SaveOutcome, vList)
                vModulesSubmitted = True
                vFDEControl.HasCompleted = True
              End If
            Case CareNetServices.FDEControlTypes.ActivityDisplay, CareNetServices.FDEControlTypes.AddressDisplay, CareNetServices.FDEControlTypes.CommunicationsDisplay, _
                 CareNetServices.FDEControlTypes.GiftAidDisplay, CareNetServices.FDEControlTypes.SuppressionDisplay
              'These modules don't actually need anything to be submitted, so just treat them as though they have been submitted
              vModulesSubmitted = True
              vFDEControl.HasCompleted = True
              vReturnDS = Nothing
          End Select
          If vReturnDS IsNot Nothing Then
            vModulesSubmitted = True
            vBatchInfo.Init(vReturnDS, vList, vBatchType)
            If mvBatchColl.ContainsKey(vBatchInfo.ToString) Then
              If vBatchInfo.CCNumber.Length > 0 AndAlso mvBatchColl(vBatchInfo.ToString).CCNumber.Length = 0 Then
                mvBatchColl(vBatchInfo.ToString).SetCCDetailsFromBatch(vBatchInfo)
              End If
            Else
              mvBatchColl.Add(vBatchInfo.ToString, vBatchInfo)
            End If
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vReturnDS)
            If vRow IsNot Nothing Then
              Dim vMsg As New StringBuilder
              If vRow.Table.Columns.Contains("PaymentPlanNumber") Then
                Dim vPPNumber As String = vRow("PaymentPlanNumber").ToString
                vMsg.AppendLine(GetInformationMessage(InformationMessages.ImPPCreated, vPPNumber))
                If vCashPP Then mvPayPlanList.Add(vPPNumber, vPPNumber)
                'Save PP numbers in a new list as mvPayPlanList gets cleared when using Next button
                mvCMDPayPlans.Add(vPPNumber, vPPNumber)
              End If
              If vRow.Table.Columns.Contains("MemberNumber") Then
                If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
                If vRow.Table.Columns.Contains("MemberNumber2") Then
                  vMsg.AppendLine(GetInformationMessage(InformationMessages.ImJointMembershipCreated, vRow("MemberNumber").ToString, vRow("MemberNumber2").ToString))
                Else
                  vMsg.AppendLine(GetInformationMessage(InformationMessages.ImMemberCreated, vRow("MemberNumber").ToString))
                End If
              End If
              If vRow.Table.Columns.Contains("DirectDebitNumber") Then
                If vMsg.ToString.Length > 0 Then vMsg.AppendLine()
                vMsg.AppendLine(GetInformationMessage(InformationMessages.ImDDCreated, vRow("DirectDebitNumber").ToString))
              End If
              If vRow.Table.Columns.Contains("Reference") Then
                vMsg.AppendLine()
                vMsg.AppendLine(GetInformationMessage(InformationMessages.ImReference, vRow("Reference").ToString))
              End If
              Dim vWarningMessage As String = ""
              If vRow.Table.Columns.Contains("WarningMessage") Then vWarningMessage = vRow("WarningMessage").ToString
              If vMsg.ToString.Length > 0 Then ShowInformationMessage(vMsg.ToString)
              If vWarningMessage.Length > 0 Then ShowWarningMessage(vWarningMessage)
              SetCMDList(vRow, vList)
            End If
          End If
        End If
        vReturnDS = Nothing
      Next

      If mvPayPlanList.Count > 0 Then
        'We have created cash Memberships, if we also added a CC donation then pay the membership
        vList = New ParameterList()
        For Each vBI As FDEBatchInfo In mvBatchColl
          If vBI.CCNumber.Length > 0 Then
            vList.IntegerValue("BatchNumber") = vBI.BatchNumber
            vList.IntegerValue("TransactionNumber") = vBI.TransactionNumber
            vList("BankAccount") = vBI.BankAccount
            Exit For
          End If
        Next
        Dim vCCNumber As String = ""
        If vList.Count = 0 Then
          'No CC donation was added - See if we just have CC details
          For Each vControl As Control In pnl.Controls
            If TypeOf (vControl) Is CareFDEControl Then
              Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
              If vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddDonationCC Or vFDEControl.ControlType = CareNetServices.FDEControlTypes.ProductSale Then
                If vFDEControl.CanSubmit = False Then
                  vFDEControl.GetPaymentMethodParameters(vList)
                  'Dont add Credit Card Number for Payment as we always want a provisional transaction
                  If vList.Contains("CreditCardNumber") Then
                    vCCNumber = vList("CreditCardNumber")
                    vList.Remove("CreditCardNumber")
                  End If
                  If vList.Count > 0 Then Exit For
                End If
              End If
            End If
          Next
        End If
        If vList.Count = 0 Then
          'No CC details are found, pay the membership by Cash
          For Each vBI As FDEBatchInfo In mvBatchColl
            If vBI.BatchType = "CA" Then
              If vBI.BatchNumber > 0 Then
                vList.IntegerValue("BatchNumber") = vBI.BatchNumber
                vList.IntegerValue("TransactionNumber") = vBI.TransactionNumber
              End If
              vList("BankAccount") = vBI.BankAccount
              Exit For
            End If
          Next
          If vList.Count = 0 Then vList("BankAccount") = ""
        End If
        If vList.Count > 0 AndAlso (vList.Contains("BankAccount") = False OrElse vList("BankAccount").Length = 0) Then
          'Now get BankAccount
          For Each vBI As FDEBatchInfo In mvBatchColl
            If vBI.BankAccount.Length > 0 Then
              vList("BankAccount") = vBI.BankAccount
              Exit For
            End If
          Next
          If vList.Contains("BankAccount") AndAlso vList("BankAccount").Length = 0 Then vList.Remove("BankAccount")
        End If
        If vList.Count > 0 Then
          vList.AddConnectionData()
          vList.AddSystemColumns()
          vList.FillFromValueList(vTransList.ValueList)
          Dim vPPIDs As New StringBuilder
          For Each vItem As String In mvPayPlanList
            vList("PaymentPlanNumber") = vItem
            vReturnDS = DataHelper.AddPaymentPlanPayment(vList)
            'Add the batch info for this payment to the collection if it does not exists
            If vCCNumber.Length > 0 Then vList("CreditCardNumber") = vCCNumber
            Dim vBatchInfo As New FDEBatchInfo
            vBatchInfo.Init(vReturnDS, vList, "CA")
            If mvBatchColl.ContainsKey(vBatchInfo.ToString) = False Then mvBatchColl.Add(vBatchInfo.ToString, vBatchInfo)
            'Remove Payment Plan from collection when pNext
            If pNext AndAlso mvPayPlanList.ContainsKey(vBatchInfo.PaymentPlanNumber.ToString) Then
              If vPPIDs.Length > 0 Then vPPIDs.Append(","c)
              vPPIDs.Append(vBatchInfo.PaymentPlanNumber)
            End If
            Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vReturnDS)
            If vRow IsNot Nothing Then
              SetCMDList(vRow, vList)
            End If
          Next
          If vPPIDs.Length > 0 Then
            For Each vPPID As String In vPPIDs.ToString.Split(","c)
              mvPayPlanList.Remove(vPPID)
            Next
          End If
        End If
      End If
    End If

    If vValid = True AndAlso pNext = False Then
      If mvBatchColl.Count > 0 Then
        'Confirm provisional transactions
        Dim vBatchIDs As New StringBuilder
        For Each vBI As FDEBatchInfo In mvBatchColl
          If vBI.BatchNumber > 0 AndAlso vBI.TransactionNumber > 0 AndAlso vBI.BatchType <> "NF" Then
            Dim vConfList As New ParameterList(True, True)
            vConfList.IntegerValue("BatchNumber") = vBI.BatchNumber
            vConfList.IntegerValue("TransactionNumber") = vBI.TransactionNumber
            vConfList.IntegerValue("ContactNumber") = vBI.ContactNumber
            vConfList.IntegerValue("AddressNumber") = vBI.AddressNumber
            If vTransList.Contains("Reference") AndAlso vTransList("Reference").ToString.Length > 0 Then
              vConfList("Reference") = vTransList("Reference")
            End If
            If vTransList.Contains("EligibleForGiftAid") Then
              vConfList("EligibleForGiftAid") = vTransList("EligibleForGiftAid")
            End If
            vConfList.AddItemIfValueSet("Mailing", vTransList.ValueIfSet("Mailing"))
            If vTransList.Contains("TransactionAmount") AndAlso DoubleValue(vTransList("TransactionAmount")) <> 0 Then vConfList("TransactionAmount") = vTransList("TransactionAmount")
            If mvMultiplePages OrElse
               (vIsCardTransaction AndAlso WebBasedCardAuthoriser.IsAvailable) Then
              vConfList("AdjustTransactionAmount") = "Y"
            End If
            Dim vReturnList As ParameterList = Nothing
            Dim vAddDonationCC As New ParameterList
            For Each vControl As Control In pnl.Controls
              If TypeOf (vControl) Is CareFDEControl Then
                Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
                If vFDEControl.ControlType = CareNetServices.FDEControlTypes.AddDonationCC Or vFDEControl.ControlType = CareNetServices.FDEControlTypes.ProductSale Then
                  vFDEControl.BuildParameterList(vAddDonationCC)
                End If
              End If
            Next
            Dim vResetProgressBar As Boolean = False
            Try
              Select Case vBI.BatchType
                Case "CA"
                  If vBI.CCNumber.Length > 0 Then
                    vBI.GetCreditCardDetails(vConfList)
                    If BooleanValue(vBI.GetAuthorisation) AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_cc_authorisation_type) = "SCXLVPCSCP" Then
                      Dim vProcess As New AsyncProcessHandler(AsyncProcessHandler.AsyncProcessHandlerTypes.ConfirmCardSaleTransaction, vConfList)
                      vResetProgressBar = True
                      prgBar.Visible = True
                      vReturnList = vProcess.GetReturnList(prgBar)
                    Else
                      vReturnList = DataHelper.ConfirmCardSaleTransaction(vConfList)
                    End If
                  ElseIf vCardNumberNotRequired Then
                    vConfList.Add("CardNumberNotRequired", "Y")
                    vReturnList = DataHelper.ConfirmCardSaleTransaction(vConfList)
                  ElseIf vIsCardTransaction AndAlso
                         WebBasedCardAuthoriser.IsAvailable Then
                    Dim form As New frmCardAuthorisation(IntegerValue(vConfList("ContactNumber")),
                                                         IntegerValue(vConfList("AddressNumber")),
                                                         "FDE Transaction",
                                                         CInt(mvTotalDonationAmount * 100),
                                                         String.Empty,
                                                         vConfList)
                    form.ShowDialog()
                    vValid = form.isAuthorised
                    If vValid Then
                      vConfList("CardNumberNotRequired") = "Y"
                      vConfList("GetAuthorisation") = "Y"
                      vReturnList = DataHelper.ConfirmCardSaleTransaction(vConfList)
                    End If
                    Else
                      If vAddDonation Then
                        If vList.Contains("PaymentMethod") AndAlso vAddPaymentPlan = False Then
                          vConfList.Add("PaymentMethod", vList("PaymentMethod"))
                          vConfList.Add("AddDonation", "Y")
                        ElseIf vAddDonationCC.Contains("PaymentMethod") Then
                          vConfList.Add("PaymentMethod", vAddDonationCC("PaymentMethod"))
                          vConfList.Add("AddDonation", "Y")
                        End If
                      End If
                      vReturnList = DataHelper.ConfirmCashSaleTransaction(vConfList)
                    End If
              End Select
            Catch vCareEx As CareException
              Select Case vCareEx.ErrorNumber
                Case CareException.ErrorNumbers.enTransAmountNotMatched, CareException.ErrorNumbers.enCCAuthorisationFailed, CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout
                  vValid = False
                  'Always remove the batch info from collection
                  If vBatchIDs.Length > 0 Then vBatchIDs.Append("|")
                  vBatchIDs.Append(vBI.ToString)
                  ShowErrorMessage(vCareEx.Message)
                Case Else
                  Throw vCareEx
              End Select
            Catch vEx As Exception
              Throw vEx
            Finally
              If vResetProgressBar Then prgBar.Visible = False
            End Try
            vModulesSubmitted = True
            If vReturnList IsNot Nothing Then
              If vReturnList.ContainsKey("BatchNumber") AndAlso vReturnList.ContainsKey("TransactionNumber") Then
                If mvMultiplePages AndAlso vReturnList.ContainsKey("Amount") AndAlso vTransList.Contains("TransactionAmount") AndAlso _
                 DoubleValue(vTransList("TransactionAmount")) <> 0 AndAlso DoubleValue(vReturnList("Amount")) <> DoubleValue(vTransList("TransactionAmount")) Then
                  ShowInformationMessage(GetInformationMessage(InformationMessages.ImTransactionAmountAdjusted, vReturnList("Amount"), vReturnList("BatchNumber"), vReturnList("TransactionNumber")))
                Else
                  ShowInformationMessage(InformationMessages.ImTransactionReference, vReturnList("BatchNumber"), vReturnList("TransactionNumber"))
                End If
              End If
            End If
          End If
        Next
        If vBatchIDs.Length > 0 Then
          'there has been a failure so we need to remove batches.
          For Each vBatchID As String In vBatchIDs.ToString.Split("|"c)
            mvBatchColl.Remove(vBatchID)
          Next
          'set HasCompleted flags to false on the fde modules so that it will process these again on re submitting page.
          For vIndex As Integer = pnl.Controls.Count - 1 To 0 Step -1
            If TypeOf (pnl.Controls(vIndex)) Is CareFDEControl Then
              Dim vFdeControl As CareFDEControl = DirectCast(pnl.Controls(vIndex), CareFDEControl)
              If vFdeControl.HasCompleted = True Then vFdeControl.HasCompleted = False
            End If
          Next
        End If
      End If
      'Now process any CMD's
      If mvCMDList.Count > 0 AndAlso vValid Then
        CreateCMD(mvCMDList)
        mvCMDFileName = ""
      End If
    End If
    If vValid = True AndAlso vModulesSubmitted = False Then
      'We didn't do anything so display Confirm Cancel message (user may have hit 'Submit' in error)
      If vTelemarketing Then
        'Give the users an option to Submit the data when they click on Next and no other module is submitted
        vValid = ShowQuestion(InformationMessages.ImConfirmDataSubmission, MessageBoxButtons.YesNo) = vbYes
        If vValid Then
          cmdOK.PerformClick()
          vValid = False
        End If
      Else
        vValid = ConfirmCancel()
      End If
    End If

    If vValid = True AndAlso pNext = False Then
      For Each vControl As Control In pnl.Controls
        If TypeOf (vControl) Is CareFDEControl Then
          Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
          vFDEControl.epl.DataChanged = False
        End If
      Next
      mvDataChanged = False
    End If

    Return vValid

  End Function

  Private Sub CMDActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFileName As String)
    mvCMDFileName = pFileName
  End Sub

  Private Sub CreateCMD(ByVal pCMDList As ParameterList)    ', ByVal pDataSet As DataSet, ByVal plist As ParameterList)
    If pCMDList.Count > 0 AndAlso pCMDList.ContainsKey("CreateMailingDocument") AndAlso pCMDList("CreateMailingDocument") = "Y" Then
      Dim vShowParagraphs As DialogResult = System.Windows.Forms.DialogResult.Yes
      If pCMDList.ContainsKey("ContactWarningSuppressionsPrompt") AndAlso pCMDList.ContainsKey("WarningSuppressions") AndAlso pCMDList("ContactWarningSuppressionsPrompt") = "Y" AndAlso pCMDList("WarningSuppressions").Length > 0 Then
        vShowParagraphs = ShowQuestion(QuestionMessages.QmWarningSuppressions, MessageBoxButtons.YesNo, pCMDList("WarningSuppressions"))
      End If
      If vShowParagraphs = System.Windows.Forms.DialogResult.Yes Then
        'Retrieve the matching paragraphs
        Dim vCMDList As New ParameterList(True)
        vCMDList.Item("Mailing") = pCMDList("Mailing")
        vCMDList.IntegerValue("ContactNumber") = pCMDList.IntegerValue("ContactNumber")
        vCMDList.Item("ExistingTransaction") = "N"
        vCMDList.Item("NewPayerContact") = "N"
        If pCMDList.ContainsKey("BatchNumber") Then vCMDList.IntegerValue("BatchNumber") = pCMDList.IntegerValue("BatchNumber")
        If pCMDList.ContainsKey("TransactionNumber") Then vCMDList.IntegerValue("TransactionNumber") = pCMDList.IntegerValue("TransactionNumber")
        If mvCMDPayPlans.Count > 0 Then
          vCMDList.Item("PaymentPlanNumber") = mvCMDPayPlans(0).ToString
          vCMDList.Item("PaymentPlanCreated") = "Y"
        End If
        If pCMDList.ContainsKey("DirectDebitNumber") Then vCMDList.Item("AutoPaymentCreated") = "Y"

        Dim vCount As Integer
        Do
          Dim vCMDDataSet As DataSet = DataHelper.GetMailingDocumentParagraphs(vCMDList)
          'Display the matching paragraphs
          Dim vParagraphsTable As DataTable = DataHelper.GetTableFromDataSet(vCMDDataSet)
          If vParagraphsTable IsNot Nothing AndAlso _
             vParagraphsTable.Columns.Contains("DisplayParagraphs") AndAlso vParagraphsTable.Rows(0).Item("DisplayParagraphs").ToString = "Y" Then
            DocumentApplication = New WordApplication
            AddHandler DocumentApplication.ActionComplete, AddressOf CMDActionComplete
            Dim vTransDocumentType As frmTransactionDocument.TransactionDocumentTypes = frmTransactionDocument.TransactionDocumentTypes.tdtTransaction
            If vCount > 0 Then vTransDocumentType = frmTransactionDocument.TransactionDocumentTypes.tdtPaymentPlan
            Dim vForm As frmTransactionDocument = New frmTransactionDocument(vTransDocumentType, vCMDDataSet, vCMDList)
            vForm.ShowDialog()
            'need to edit the document when Edit is pressed
            vCMDDataSet = vForm.DataSet
          End If
          'Create the mailing document
          If vCMDList.Contains("EarliestFulfilmentDate") Then
            If mvCMDFileName.Length = 0 Then
              Dim vSelectedParagraphs As New StringBuilder
              Dim vCMDTable As DataTable = DataHelper.GetTableFromDataSet(vCMDDataSet)
              If vCMDTable IsNot Nothing Then
                For Each vCMDRow As DataRow In vCMDTable.Rows
                  If CBool(vCMDRow.Item("Include")) Then
                    If vSelectedParagraphs.Length > 0 Then vSelectedParagraphs.Append(",")
                    vSelectedParagraphs.Append(vCMDRow.Item("ParagraphNumber"))
                  End If
                Next
              End If
              If vSelectedParagraphs.Length = 0 Then
                vCMDList("SelectedParagraphs") = "0"
              Else
                vCMDList("SelectedParagraphs") = vSelectedParagraphs.ToString
              End If
            End If
            vCMDDataSet = DataHelper.AddContactMailingDocument(vCMDList)
            If mvCMDFileName.Length > 0 Then
              Dim vResultRow As DataRow = vCMDDataSet.Tables("Result").Rows(0)
              DataHelper.UpdateContactMailingDocumentFile(IntegerValue(vResultRow.Item("MailingDocumentNumber").ToString), mvCMDFileName)
            End If

            vCount += 1
            If mvCMDPayPlans.Count > 0 AndAlso mvCMDPayPlans.Count > vCount Then
              With vCMDList
                .Item("PaymentPlanNumber") = mvCMDPayPlans(vCount)
                If .Contains("BatchNumber") Then .Remove("BatchNumber")
                If .Contains("TransactionNumber") Then .Remove("TransactionNumber")
                If .Contains("DeclarationNumber") Then .Remove("DeclarationNumber")
              End With
            End If
          End If
        Loop While vCount < mvCMDPayPlans.Count
      End If
    End If
  End Sub

#End Region

#Region " User Control Events "

  Private Sub CareFDEControl_ContactChanged(ByVal sender As Object, ByVal pContactNumber As Integer)
    If pContactNumber > 0 Then
      mvContactInfo = New ContactInfo(pContactNumber)
      If mvContactInfo.ContactNumber > 0 Then
        For Each vControl As Control In pnl.Controls
          If TypeOf (vControl) Is CareFDEControl Then
            Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
            If vFDEControl.SupportsContactData Then
              If mvEditMode <> EditMode.Edit Then vFDEControl.RefreshContactData(mvContactInfo)
            End If
          End If
        Next
      End If
    End If
  End Sub

  Private Sub CareFDEControl_ReferenceMandatory(ByVal sender As Object, ByVal pMandatory As Boolean)
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If vFDEControl.SupportsReferenceData = True Then
          vFDEControl.SetReferenceMandatory(pMandatory)
        End If
      End If
    Next
  End Sub

  Private Sub CareFDEControl_RegDonationAdded(ByVal sender As Object)
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is FDEAddRegularDonation Then
        Dim vFDEControl As FDEAddRegularDonation = DirectCast(vControl, FDEAddRegularDonation)

        '    vFDEControl.RefreshDonationBalance()

      End If
    Next
  End Sub
  Private Sub CareFDEControl_ClearBankDetails(ByVal sender As Object)
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If vFDEControl.SupportsClearBankDetails = True Then
          vFDEControl.ResetBankDetails()
        End If
      End If
    Next
  End Sub

  Private Sub CareFDEControl_SelectedContactChanged(ByVal sender As Object, ByVal pContactNumber As Integer)
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If vFDEControl.SupportsSelectionChanged Then
          vFDEControl.ChangeSelectedContact(pContactNumber)
        End If
      End If
    Next
  End Sub

  Private Sub CareFDEControl_AddressChanged(ByVal sender As Object, ByVal pAddressNumber As Integer)
    If pAddressNumber > 0 Then
      For Each vControl As Control In pnl.Controls
        If TypeOf (vControl) Is CareFDEControl Then
          Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
          If vFDEControl.SupportsAddressData Then
            If mvEditMode <> EditMode.Edit Then vFDEControl.RefreshAddressData(pAddressNumber)
          End If
        End If
      Next
    End If
  End Sub

  Private Sub CareFDEControl_SourceChanged(ByVal sender As Object, ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is CareFDEControl Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If vFDEControl.SupportsSourceChanged Then
          If mvEditMode <> EditMode.Edit Then vFDEControl.RefreshSource(pSourceCode, pDistributionCode, pIncentiveScheme)
        End If
      End If
    Next
  End Sub

  Private Sub CareFDEControl_SetAndRefreshSource(ByVal sender As Object, ByVal pSourceCode As String, ByVal pDistributionCode As String, ByVal pIncentiveScheme As String)
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is FDEAddTransactionDetails Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If mvEditMode <> EditMode.Edit Then
          vFDEControl.SetAndRefreshSource(pSourceCode)
          Exit For
        End If
      End If
    Next
  End Sub

  Private Sub careFDEControl_TransactionDateChanged(ByVal sender As Object, ByVal pTransactionDate As String)
    For Each vControl As Control In pnl.Controls
      If TypeOf vControl Is CareFDEControl Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If vFDEControl.SupportsTransactionDateChanged Then
          If mvEditMode <> EditMode.Edit Then vFDEControl.RefreshTransactionDate(pTransactionDate)
        End If
      End If
    Next
  End Sub

  Private Sub CareFDEControl_DeleteControl(ByVal sender As Object)
    Dim vFDEControl As CareFDEControl = DirectCast(sender, CareFDEControl)
    vFDEControl.Delete()
    pnl.Visible = False                   'Stop screen flicker
    pnl.Controls.Remove(vFDEControl)
    pnl.Visible = True
  End Sub

  Private Sub CareFDEControl_EnableOtherModules(ByVal sender As Object, ByVal pEnable As Boolean)
    Dim vSender As Control = CType(sender, Control)
    For Each vControl As Control In pnl.Controls
      If vControl IsNot vSender Then
        Dim vFDEControl As CareFDEControl = DirectCast(vControl, CareFDEControl)
        If mvEditMode <> EditMode.Edit AndAlso Not (TypeOf vSender Is FDETelemarketing AndAlso TypeOf vFDEControl Is FDEContactSelection AndAlso pEnable) Then vFDEControl.Enabled = pEnable
      End If
    Next
    cmdNext.Enabled = pEnable
    cmdOK.Enabled = pEnable
  End Sub

  Private Sub CareFDEControl_FormClosingAllowed(ByVal pValue As Boolean)
    mvFormClosingAllowed = pValue
    cmdCancel.Enabled = pValue
  End Sub
#End Region

#Region " Buttons "

  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      If mvEditMode = EditMode.Run Then
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Else
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        If mvEditMode = EditMode.Edit Then
          SaveFDEPage()
          SaveFDEPageItems()
          mvDataChanged = False
        End If
      End If
      Close()
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    'This is the Submit button
    Dim vBusyCursor As New BusyCursor()
    Try
      mvCMDFileName = ""
      If mvEditMode = EditMode.Run Then
        If ProcessSubmit(False) Then
          Dim vFDEControl As CareFDEControl = Nothing
          Dim vTelemarketingSource As String = ""
          For Each vControl As Control In pnl.Controls
            If TypeOf (vControl) Is CareFDEControl Then
              vFDEControl = DirectCast(vControl, CareFDEControl)
              vFDEControl.ResetModule(False)
              If TypeOf vFDEControl Is FDETelemarketing Then vTelemarketingSource = vFDEControl.epl.FindPanelControl(Of TextLookupBox)("Segment").GetDataRowItem("Source")
              If TypeOf vFDEControl Is FDEAddTransactionDetails AndAlso vTelemarketingSource.Length > 0 Then vFDEControl.SetAndRefreshSource(vTelemarketingSource)
            End If
          Next
          mvDataChanged = False
          CareFDEControl_FormClosingAllowed(True)
          mvBatchColl = New CollectionList(Of FDEBatchInfo)
          mvPayPlanList = New CollectionList(Of String)
          mvCMDPayPlans = New CollectionList(Of String)
          mvMultiplePages = False
          mvCMDList = New ParameterList
          mvTotalDonationAmount = 0
          mvPreviousCardNumberNotRequired = False
        End If
      End If
      Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Catch vCareEX As CareException
      Select Case vCareEX.ErrorNumber
        Case CareException.ErrorNumbers.enNoTransactionDetails, CareException.ErrorNumbers.enNoContactSelected
          ShowErrorMessage(vCareEX.Message)
        Case Else
          DataHelper.HandleException(vCareEX)
      End Select
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      mvBatchNumber = 0
      mvTransactionNumber = 0
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim vBusyCursor As New BusyCursor()
    Try
      If mvEditMode = EditMode.Edit Then
        SaveFDEPage()
        SaveFDEPageItems()
        mvDataChanged = False
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdTest_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdTest.Click
    Try
      If mvEditMode = EditMode.Edit Then
        SaveFDEPage()
        SaveFDEPageItems()
        mvDataChanged = False
      End If
      Dim vForm As New frmFastDataEntry(mvPageNumber, True)
      vForm.ShowDialog(Me)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub cmdNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNext.Click
    Dim vBusyCursor As New BusyCursor()
    Try
      If mvCMDFileName Is Nothing Then mvCMDFileName = ""
      If ProcessSubmit(True) Then
        For Each vControl As Control In pnl.Controls
          If TypeOf (vControl) Is CareFDEControl Then
            DirectCast(vControl, CareFDEControl).ResetModule(True)
          End If
        Next
      End If
    Catch vCareEX As CareException
      Select Case vCareEX.ErrorNumber
        Case CareException.ErrorNumbers.enNoTransactionDetails, CareException.ErrorNumbers.enNoContactSelected
          ShowErrorMessage(vCareEX.Message)
        Case Else
          DataHelper.HandleException(vCareEX)
      End Select
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Sub SetCMDList(ByVal pRow As DataRow, ByVal pList As ParameterList)
    If mvCMDList.Count = 0 AndAlso pRow.Table.Columns.Contains("CreateMailingDocument") AndAlso pRow.Item("CreateMailingDocument").ToString = "Y" Then
      mvCMDList("CreateMailingDocument") = "Y"
      If pRow.Table.Columns.Contains("ContactWarningSuppressionsPrompt") AndAlso pRow.Table.Columns.Contains("ContactWarningSuppressions") AndAlso pRow.Item("ContactWarningSuppressionsPrompt").ToString = "Y" AndAlso pRow.Item("ContactWarningSuppressions").ToString.Length > 0 Then
        mvCMDList("ContactWarningSuppressionsPrompt") = "Y"
        mvCMDList("WarningSuppressions") = pRow.Item("ContactWarningSuppressions").ToString
      End If
      If pList.ContainsKey("Mailing") Then
        mvCMDList("Mailing") = pList("Mailing")
      ElseIf pRow.Table.Columns.Contains("Mailing") Then
        mvCMDList("Mailing") = pRow.Item("Mailing").ToString
      End If
      mvCMDList.IntegerValue("ContactNumber") = pList.IntegerValue("PayerContactNumber")
    End If
    If mvCMDList.Count > 0 AndAlso pRow.Table.Columns.Contains("CreateMailingDocument") AndAlso pRow.Item("CreateMailingDocument").ToString = "Y" Then
      If pRow.Table.Columns.Contains("BatchNumber") AndAlso IntegerValue(pRow.Item("BatchNumber").ToString) > 0 Then mvCMDList.IntegerValue("BatchNumber") = IntegerValue(pRow.Item("BatchNumber").ToString)
      If pRow.Table.Columns.Contains("TransactionNumber") AndAlso IntegerValue(pRow.Item("TransactionNumber").ToString) > 0 Then mvCMDList.IntegerValue("TransactionNumber") = IntegerValue(pRow.Item("TransactionNumber").ToString)
      If pRow.Table.Columns.Contains("PaymentPlanNumber") Then mvCMDList.IntegerValue("PaymentPlanNumber") = IntegerValue(pRow.Item("PaymentPlanNumber").ToString)
      If pRow.Table.Columns.Contains("DirectDebitNumber") Then mvCMDList.IntegerValue("DirectDebitNumber") = IntegerValue(pRow.Item("DirectDebitNumber").ToString)
    End If
  End Sub
#End Region

#Region " FDEBatchInfo Class "

  Private Class FDEBatchInfo

    Public BatchNumber As Integer
    Public TransactionNumber As Integer
    Public BatchType As String = ""
    Public ContactNumber As Integer
    Public AddressNumber As Integer
    Public CCNumber As String = ""
    Public CardStartDate As String = ""
    Public CardExpiryDate As String = ""
    Public CreditCardType As String = ""
    Public GetAuthorisation As String = ""
    Public BankAccount As String = ""
    Public PaymentPlanNumber As Integer

    Friend Sub Init(ByVal pDS As DataSet, ByVal pList As ParameterList, ByVal pBatchType As String)
      Dim vDT As DataTable = DataHelper.GetTableFromDataSet(pDS)
      If vDT.Rows.Count > 0 AndAlso vDT.Columns.Contains("BatchNumber") Then
        BatchNumber = IntegerValue(vDT.Rows(0).Item("BatchNumber").ToString)
        TransactionNumber = IntegerValue(vDT.Rows(0).Item("TransactionNumber").ToString)
      End If
      BatchType = pBatchType
      If pBatchType <> "NF" AndAlso pList.ContainsKey("CreditCardNumber") Then
        CCNumber = pList("CreditCardNumber")
        If pList.ContainsKey("CardStartDate") Then CardStartDate = pList("CardStartDate")
        CardExpiryDate = pList("CardExpiryDate")
        CreditCardType = pList("CreditCardType")
        If pList.ContainsKey("GetAuthorisation") Then GetAuthorisation = pList("GetAuthorisation")
      End If
      ContactNumber = pList.IntegerValue("PayerContactNumber")
      AddressNumber = pList.IntegerValue("PayerAddressNumber")
      If pList.ContainsKey("BankAccount") Then BankAccount = pList("BankAccount")
      If pList.Contains("PaymentPlanNumber") Then PaymentPlanNumber = pList.IntegerValue("PaymentPlanNumber")
    End Sub

    Public Overrides Function ToString() As String
      Return BatchNumber.ToString & "," & TransactionNumber.ToString
    End Function

    Friend Sub GetCreditCardDetails(ByVal pList As ParameterList)
      If CCNumber.Length > 0 Then
        pList("CreditCardNumber") = CCNumber
        pList("CardStartDate") = CardStartDate
        pList("CardExpiryDate") = CardExpiryDate
        pList("CreditCardType") = CreditCardType
        If GetAuthorisation.Length > 0 Then pList("GetAuthorisation") = GetAuthorisation
      End If
    End Sub

    Friend Sub SetCCDetailsFromBatch(ByVal pBatchInfo As FDEBatchInfo)
      With pBatchInfo
        If .CCNumber.Length > 0 Then
          CCNumber = .CCNumber
          CardStartDate = .CardStartDate
          CardExpiryDate = .CardExpiryDate
          CreditCardType = .CreditCardType
          GetAuthorisation = .GetAuthorisation
        End If
      End With
    End Sub
  End Class

#End Region
End Class

