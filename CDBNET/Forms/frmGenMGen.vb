Imports CDBNET.MailingInfo
Imports CDBNET.GeneralMailing

Public Class frmGenMGen

  Private mvIncludeTYLs As Boolean
  Private Enum MailingType
    GeneralMailing
    DirectDebit
    StandingOrders
    MembershipCards
    SelectionTester
    StandingOrderCancellation
    GAYECancellation
  End Enum

  Private Const SORT_BY_BRANCH As Integer = 0
  Private Const SORT_BY_COUNTRY As Integer = 1
  Private Const SORT_BY_MAILSORT As Integer = 2
  Private Const SORT_BY_SURNAME As Integer = 3
  Private Const SORT_BY_OTHER_1 As Integer = 4
  Private Const SORT_BY_OTHER_2 As Integer = 5

  Private Const DEF_SORT_ORDER As Integer = SORT_BY_SURNAME

  Private mvMailingTypeCode As String = String.Empty
  Private mvCaption As String = ""
  Private mvUnloadOK As Boolean
  Private mvContactNumber As Integer
  Private mvAddressNumber As Integer
  Public mvCurrentSelectionSet As Integer       'The last set selected from GEMMLIST
  Public mvMailingInfo As MailingInfo
  Private mvSelectionSet As Integer
  Private mvReportOption As ReportOptions
  Private mvSingleMembershipCardMailing As Boolean
  Private mvOmitted As List(Of Integer)

  Public Sub New()

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.    
    InitialiseControls()
  End Sub

  Public Sub New(ByVal pTypeCode As String, ByRef pMailingInfo As MailingInfo, ByVal pSelectionSet As Integer, ByVal pSingleMembership As Boolean)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    mvMailingTypeCode = pTypeCode
    mvMailingInfo = pMailingInfo
    mvSingleMembershipCardMailing = pSingleMembership
    InitialiseControls()
    mvSelectionSet = pSelectionSet
  End Sub

  Public Sub New(ByVal pTypeCode As String, ByRef pMailingInfo As MailingInfo, ByVal pSelectionSet As Integer)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    mvMailingTypeCode = pTypeCode
    mvMailingInfo = pMailingInfo
    InitialiseControls()
    mvSelectionSet = pSelectionSet
  End Sub
  Private Sub InitialiseControls()
    SetControlTheme()
    ButtonPanel1.RepositionButtons()
    Dim vOptionalMailingHistory As Boolean

    Dim vPanelItems As New PanelItems("epl")
    Dim mvTmpDataSet As New DataSet
    Dim vSortOrderCombo As ComboBox = Nothing
    eplSOCancellation.Init((New EditPanelInfo(CDBNETCL.CareNetServices.FunctionParameterTypes.fptSOCancellation)))

    eplStandard.Init(New EditPanelInfo(If(mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyExamCertificates,
                                          CDBNETCL.CareNetServices.FunctionParameterTypes.fptExamCertificates,
                                          CDBNETCL.CareNetServices.FunctionParameterTypes.fptStandardFields)))
    eplMisc.Init(New EditPanelInfo(CDBNETCL.CareNetServices.FunctionParameterTypes.fptMiscFields))
    eplSelectionTester.Init(New EditPanelInfo(CDBNETCL.CareNetServices.FunctionParameterTypes.fptSelectionTester))

    If eplStandard.PanelInfo.PanelItems.Exists("SortOrder") Then vSortOrderCombo = eplStandard.FindComboBox("SortOrder")

    vOptionalMailingHistory = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciOptionalMailingHistory)
    mvIncludeTYLs = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.client_mc_include_TYLs)

    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGeneralMailing Then cmdReport.Visible = True
    Me.Text = mvMailingInfo.Caption + ControlText.FrmGeneralMailing '& LoadString(28112)    ' - Generate

    mvUnloadOK = False
    eplMisc.FindCheckBox("CheckBoxOne").Visible = False
    eplMisc.PanelInfo.PanelItems("CheckBoxOne").Hidden = True
    eplMisc.FindLabel("CheckBoxOne").Visible = False
    eplMisc.FindCheckBox("CheckBoxTwo").Visible = False
    eplMisc.PanelInfo.PanelItems("CheckBoxTwo").Hidden = True
    eplMisc.FindLabel("CheckBoxTwo").Visible = False
    HideMembershipsGroup(False)
    eplMisc.Visible = False
    eplSelectionTester.Visible = False
    pnlSOCancellation.Visible = False
    drg.MaxGridRows = DisplayTheme.DefaultMaxGridRows

    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
      eplStandard.FindCheckBox("CreateMailingHistory").Visible = vOptionalMailingHistory
      eplStandard.PanelInfo.PanelItems("CreateMailingHistory").Hidden = Not vOptionalMailingHistory
      eplStandard.FindLabel("CreateMailingHistory").Visible = vOptionalMailingHistory
      eplStandard.FindTextBox("MailingNotes").Visible = True
      eplStandard.PanelInfo.PanelItems("MailingNotes").Hidden = False
    End If

    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyMembershipCards Then
      If vSortOrderCombo IsNot Nothing Then
        With vSortOrderCombo
          .Items.Add(ControlText.CboSortOrderBranch)
          .Items.Add(ControlText.CboSortOrderCountry)
          .Items.Add(ControlText.CboSortOrderMailingSort)
          .Items.Add(ControlText.CboSortOrderSurname)
        End With
      End If
    End If

    If mvSingleMembershipCardMailing Then
      cmdClear.Visible = False
      cmdCount.Visible = False
      cmdFindAddress.Visible = False
      cmdMerge.Visible = False
      cmdOmit.Visible = False
      cmdPrint.Visible = False
      cmdRefine.Visible = False
      cmdReset.Visible = False
      cmdSaveCriteria.Visible = False
      cmdSaveList.Visible = False
      cmdView.Visible = False
      pnlAddress.Visible = False
      pnlGrid.Visible = False
      SplitContainer1.SplitterDistance = 0
      SplitContainer1.Panel2Collapsed = True
      Me.Height = 431
    End If

    'populate standard document field (for all mailing types)
    Dim vList As New ParameterList(True)
    vList("ApplicationName") = mvMailingTypeCode
    Dim vDataTable As New DataTable
    vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtReports, vList)
    Dim vMailmergeHeader As String = ""
    If vDataTable IsNot Nothing Then
      For Each vRow As DataRow In vDataTable.Rows
        vMailmergeHeader = vRow("Mailmergeheader").ToString
      Next
    End If
    Dim vPanelItem As PanelItem = Nothing
    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
      eplStandard.FindTextLookupBox("StandardDocument").FillComboWithRestriction(vMailmergeHeader)
      eplStandard.FindTextLookupBox("StandardDocument").Visible = True
      eplStandard.PanelInfo.PanelItems("StandardDocument").Hidden = False
      eplStandard.SetValue("StandardDocument", "", False, False, True, True) 'BR16311/16322 standard document textlookupbox should not have a value selected if the default value is not set and a document exists
      vPanelItem = DirectCast(eplStandard.FindTextLookupBox("StandardDocument").Tag, PanelItem)
      vPanelItem.ValueChanged("")   'BR16319 need to call value changed so that the initial value is set correctly in the panel item object
    End If

    Select Case mvMailingInfo.MailingType
      Case CareNetServices.MailingTypes.mtyDirectDebits
        eplStandard.Dock = DockStyle.Fill
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.Items.Add(ControlText.CboSortOrderDirectDebit) 'Direct Debit Number
      Case CareNetServices.MailingTypes.mtyMembers
        eplMisc.Visible = True
        eplMisc.Dock = DockStyle.Bottom
        eplStandard.Dock = DockStyle.Top

        eplMisc.FindCheckBox("CheckBoxOne").Visible = True  ' pnlCheckBoxes.Visible = True
        eplMisc.PanelInfo.PanelItems("CheckBoxOne").Hidden = False
        eplMisc.FindCheckBox("CheckBoxTwo").Visible = False 'chk1.Visible = False 'HIDE THIS UNTIL THE CORP. MEM'SHIP REPORT IS WRITTEN
        eplMisc.PanelInfo.PanelItems("CheckBoxTwo").Hidden = True
        eplMisc.FindLabel("CheckBoxTwo").Visible = False
        eplMisc.FindCheckBox("CheckBoxOne").Text = ControlText.ChkCombineMail   'Combine Mail - Same Address && Payment Plan
        eplMisc.FindLabel("CheckBoxOne").Text = ControlText.ChkCombineMail   'Combine Mail - Same Address && Payment Plan
        eplMisc.FindCheckBox("CheckBoxTwo").Text = ControlText.ChkCorporateMember    'Corporate Memberships
        eplMisc.FindLabel("CheckBoxTwo").Text = ControlText.ChkCorporateMember    'Corporate Memberships
        If vSortOrderCombo IsNot Nothing Then
          vSortOrderCombo.Items.Add(ControlText.CboSortOrderMemberNumber)    'Member Number
          vSortOrderCombo.Items.Add(ControlText.CboSortOrderMembershipType)    'Membership Type
        End If
      Case CareNetServices.MailingTypes.mtyMembershipCards
        eplMisc.Visible = True
        eplMisc.Dock = DockStyle.Bottom
        eplStandard.Dock = DockStyle.Top
        HideMembershipsGroup(True)
        eplMisc.FindCheckBox("DeleteThankYou").Enabled = False
        eplMisc.FindTextLookupBox("Company").ComboBox.SelectedIndex = 0
        eplMisc.FindTextLookupBox("Company").Enabled = False
        eplMisc.FindTextLookupBox("Mailing").Enabled = False

        eplMisc.FindCheckBox("CheckBoxOne").Text = ControlText.ChkCorporateMember    'Corporate Memberships
        eplMisc.FindCheckBox("CheckBoxTwo").Visible = False
        eplMisc.PanelInfo.PanelItems("CheckBoxTwo").Hidden = True
        eplMisc.FindLabel("CheckBoxTwo").Visible = False
        If vSortOrderCombo IsNot Nothing Then
          vSortOrderCombo.Items.Add(ControlText.CboSortOrderMemberNumber)    'Member Number
          vSortOrderCombo.Items.Add(ControlText.CboSortOrderGiftAndMember)    'Membership Type, Gift & Member Number
          vSortOrderCombo.Items.Add(ControlText.CboSortOrderMailingSort)    'Mailsort
          vSortOrderCombo.SelectedIndex = SORT_BY_BRANCH
        End If

        'If eplMisc.Visible Then
        '  eplMisc.FindTextLookupBox("Company").ComboBox.SelectedIndex = 0
        '  eplMisc.FindTextLookupBox("Company").Enabled = False
        '  eplMisc.FindLabel("Company").Enabled = False
        'End If

        If AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.card_default_standard_document).Length > 0 AndAlso mvMailingTypeCode = "MC" Then
          Dim vText As String = AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.card_default_standard_document)
          eplStandard.FindTextLookupBox("StandardDocument").Text = vText
          vPanelItem.ValueChanged(vText)  'BR16319 need to call value changed so that the initial value is set correctly in the panel item object (mvLastValue)
          eplStandard.FindCheckBox("AutoMerge").Enabled = True
        End If

      Case CareNetServices.MailingTypes.mtyPayers
        ShowMiscCheckBoxes(True)  'pnlCheckBoxes.Visible = True 'HIDE THIS UNTIL THE CORP. REPORT IS WRITTEN
        eplMisc.FindCheckBox("CheckBoxOne").Text = ControlText.ChkVatAndGroup      'VAT && Group Details
        eplMisc.FindCheckBox("CheckBoxTwo").Visible = False
        eplMisc.PanelInfo.PanelItems("CheckBoxTwo").Hidden = True
        eplMisc.FindLabel("CheckBoxTwo").Visible = False
        eplStandard.Dock = DockStyle.Fill
      Case CareNetServices.MailingTypes.mtySelectionTester
        eplStandard.Visible = False
        pnlGrid.Visible = False
        pnlAddress.Visible = False
        eplSelectionTester.Visible = True
        SplitContainer1.Panel2Collapsed = True
        SplitContainer1.Dock = DockStyle.Fill
        eplSelectionTester.Dock = DockStyle.Fill
        pnlEpl.Dock = DockStyle.Fill
      Case CareNetServices.MailingTypes.mtyStandingOrders
        eplStandard.Dock = DockStyle.Fill
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.Items.Add(ControlText.CboSortOrderStandingOrderNumber) 'Standing Order Number
      Case CareNetServices.MailingTypes.mtySubscriptions
        eplMisc.Visible = True
        eplMisc.Dock = DockStyle.Bottom
        eplStandard.Dock = DockStyle.Top
        ShowMiscCheckBoxes(True)
        eplMisc.FindLabel("CheckBoxOne").Text = ControlText.ChkCombineMail    'Combine Mail - Same Address && Payment Plan
        eplMisc.FindLabel("CheckBoxTwo").Text = ControlText.ChkCombineMailSameAddress    'Combine Mail - Same Address Only
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.Items.Add(ControlText.CboSortOrderDespatchMethod) 'Despatch Method
      Case CareNetServices.MailingTypes.mtyStandingOrderCancellation, CareNetServices.MailingTypes.mtyGAYECancellation
        eplStandard.Visible = False
        eplSOCancellation.Visible = True
        pnlSOCancellation.Visible = True
        eplSOCancellation.Dock = DockStyle.Fill
        pnlEpl.Visible = False
      Case CareNetServices.MailingTypes.mtyMemberFulfilment
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.Items.Add(ControlText.CboSortOrderNumber) 'Order Number
        eplMisc.Visible = True  'pnlMisc.Visible = True
        eplMisc.Dock = DockStyle.Bottom
        eplStandard.Dock = DockStyle.Top
        ShowMiscCheckBoxes(True)  'pnlCheckBoxes.Visible = True
        eplMisc.FindLabel("CheckBoxOne").Text = ControlText.CboSortOrderMailMembership
        eplMisc.FindCheckBox("CheckBoxTwo").Visible = False
        eplMisc.PanelInfo.PanelItems("CheckBoxTwo").Hidden = True
        eplMisc.FindLabel("CheckBoxTwo").Visible = False
      Case CareNetServices.MailingTypes.mtyNonMemberFulfilment
        eplStandard.Dock = DockStyle.Fill
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.Items.Add(ControlText.CboSortOrderNumber) 'Order Number
      Case CareNetServices.MailingTypes.mtyExamCertificates
        eplSOCancellation.Visible = False
        eplMisc.Visible = False
        eplSelectionTester.Visible = False
        Dim vExamSelector As ExamSelector = TryCast(FindControl(eplStandard, "ExamUnitLinkId"), ExamSelector)
        If vExamSelector IsNot Nothing Then
          vExamSelector.Init(ExamSelector.SelectionType.Courses)
          Dim vTree As TreeView = TryCast(FindControl(vExamSelector, "tvw"), TreeView)
          Dim vCertRunType As ComboBox = eplStandard.FindPanelControl(Of ComboBox)("ExamCertRunType")
          HideNonUnitNodes(vTree.Nodes)
          AddHandler vExamSelector.ItemSelected, Sub(sender As Object,
                                                     pSelectionType As ExamsAccess.XMLExamDataSelectionTypes,
                                                     pSelectionItem As ExamSelectorItem)
                                                   drg.Clear()
                                                   lblAddress.Text = ""
                                                   lblOrganisation.Text = ""
                                                   cmdOmit.Enabled = False
                                                   cmdFindAddress.Enabled = False
                                                   Dim vTempList As New ParameterList(True, True)
                                                   vTempList.Add("ExamUnitLinkId", DirectCast(sender, ExamSelector).GetUnitLinkID)
                                                   Dim vRunTypes As DataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamUnitCertRunTypes, vTempList)
                                                   If vRunTypes.Tables.Contains("DataRow") Then
                                                     vCertRunType.DataSource = vRunTypes.Tables("DataRow")
                                                     vCertRunType.DisplayMember = "ExamCertRunTypeDesc"
                                                     vCertRunType.ValueMember = "ExamCertRunType"
                                                     cmdView.Enabled = True
                                                   Else
                                                     vCertRunType.DataSource = Nothing
                                                     cmdView.Enabled = False
                                                   End If
                                                 End Sub
          AddHandler vCertRunType.SelectedValueChanged, Sub(sender As Object, e As EventArgs)
                                                          CheckMandatory()
                                                          cmdView.Enabled = DirectCast(sender, ComboBox).SelectedItem IsNot Nothing
                                                        End Sub
          drg.Clear()
          vTree.SelectedNode = vTree.Nodes(0)
        End If
    End Select
    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyMembershipCards Then
      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandingOrderCancellation Or mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGAYECancellation Then
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.SelectedIndex = -1
      Else
        If vSortOrderCombo IsNot Nothing Then vSortOrderCombo.SelectedIndex = DEF_SORT_ORDER
      End If
    End If

    eplStandard.FindComboBox("MailMerge").Items.Add(ControlText.CboMailMergerNone) '<None>
    eplStandard.FindComboBox("MailMerge").Items.Add(ControlText.CboMailMergeWord)
    eplStandard.FindComboBox("MailMerge").SelectedIndex = 1

    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates AndAlso
       eplStandard.FindTextLookupBox("StandardDocument").Text = "" Then
      eplStandard.FindCheckBox("AutoMerge").Enabled = False 'set auto merge to disabled if no standard doc selected
    End If

    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
      eplStandard.FindTextLookupBox("Device").ComboBox.SelectedIndex = 0
    End If

    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_auto_name_mailing_files, False) Then
      eplStandard.FindTextBox("ReportDestination").Text = AppValues.GetMailingFileName() 'DataHelper.GetMailingFile()   gvEnv.GetMailingFileName(True, 0) pDataRow("MailingFilename").ToString
      eplStandard.FindTextBox("ReportDestination").Enabled = False
      eplStandard.FindButton("Browse").Enabled = False
    End If

    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
      eplStandard.FindDateTimePicker("Date").Value = CDate(AppValues.TodaysDate)
      eplStandard.FindDateTimePicker("Date").MinDate = CDate(AppValues.TodaysDate)    '    dtpMailingDate.Value = TodaysDate
      eplStandard.FindDateTimePicker("Date").ShowCheckBox = False
    End If

    cmdOmit.Enabled = False
    cmdFindAddress.Enabled = False
    cmdSaveCriteria.Enabled = CBool(IIf((mvMailingInfo.CriteriaRows > 0), True, False))
    cmdMerge.Enabled = mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGeneralMailing
    cmdCount.Visible = mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySubscriptions
    ''When all of the summary reports have been written and we re-enable the Print
    ''button make sure that similar code is removed from the cboSortOrder_Click event
    cmdPrint.Enabled = mvMailingInfo.SummaryPrintValid
    cmdView.Enabled = Not mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySelectionTester
    cmdRefine.Enabled = Not mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySelectionTester
    cmdReset.Enabled = Not mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySelectionTester
    CheckMandatory()

    ' Print button is disabled for Irish Gift Aid as this functionality is not working in Rich Client also
    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyIrishGiftAid Then
      cmdPrint.Enabled = False
    End If

    eplStandard.AdjustVisibleControls()
    eplMisc.AdjustVisibleControls()
    eplSelectionTester.AdjustVisibleControls()
    eplSOCancellation.AdjustVisibleControls()
    eplSOCancellation.AdjustVisibleControls()

    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyExamCertificates Then
      cmdView.Enabled = eplStandard.FindPanelControl(Of ComboBox)("ExamCertRunType").SelectedItem IsNot Nothing
    End If

  End Sub

  Private Sub HideNonUnitNodes(pNodes As TreeNodeCollection)
    If pNodes IsNot Nothing Then
      Dim vUnwantedNodes As New List(Of TreeNode)
      For Each vNode As TreeNode In pNodes
        If vNode.Tag IsNot Nothing Then
          Dim vExamSelectorItem As ExamSelectorItem = TryCast(vNode.Tag, ExamSelectorItem)
          If vExamSelectorItem IsNot Nothing Then
            If vExamSelectorItem.ExamSelectionType <> ExamsAccess.XMLExamDataSelectionTypes.ExamUnits Then
              vUnwantedNodes.Add(vNode)
            Else
              HideNonUnitNodes(vNode.Nodes)
            End If
          End If
        End If
      Next vNode
      For Each vNode As TreeNode In vUnwantedNodes
        pNodes.Remove(vNode)
      Next vNode
    End If
  End Sub

  Private Sub HideMembershipsGroup(ByVal pVisible As Boolean)
    eplMisc.FindTextLookupBox("Mailing").Visible = pVisible
    eplMisc.FindTextLookupBox("Mailing").Enabled = pVisible
    eplMisc.PanelInfo.PanelItems("Mailing").Hidden = Not pVisible
    eplMisc.FindLabel("Mailing").Visible = pVisible
    eplMisc.FindCheckBox("IncludeThankYou").Visible = pVisible
    eplMisc.FindCheckBox("IncludeThankYou").Enabled = pVisible
    eplMisc.PanelInfo.PanelItems("IncludeThankYou").Hidden = Not pVisible
    eplMisc.FindLabel("IncludeThankYou").Visible = pVisible
    eplMisc.FindCheckBox("DeleteThankYou").Visible = pVisible
    eplMisc.FindCheckBox("DeleteThankYou").Enabled = pVisible
    eplMisc.PanelInfo.PanelItems("DeleteThankYou").Hidden = Not pVisible
    eplMisc.FindLabel("DeleteThankYou").Visible = pVisible
    eplMisc.FindTextLookupBox("Company").Visible = pVisible
    eplMisc.FindTextLookupBox("Company").Enabled = pVisible
    eplMisc.PanelInfo.PanelItems("Company").Hidden = Not pVisible
    eplMisc.FindLabel("Company").Visible = pVisible
  End Sub

  Private Sub eplStandard_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles eplStandard.ValueChanged
    Try
      Select Case pParameterName
        Case "MailMerge", "ReportDestination"
          CheckMandatory()
        Case "Mailing"
          If pValue.Length > 0 Then
            Dim vList As New ParameterList(True, True)
            vList("Mailing") = pValue
            Dim vRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMailings, vList)
            If vRow Is Nothing Then
              'Create a new Mailing code
              Try
                Dim vDefaults As New ParameterList
                vDefaults("Mailing") = pValue
                Dim vForm As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptNewMailingCode, vDefaults, Nothing)
                If vForm.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                  Dim vResult As ParameterList = DataHelper.AddLookupData("Mailings", vForm.ReturnList)
                  eplStandard.SetValue("Mailing", vResult("Mailing").ToString)   'Use the MailingCode from the results in case the user changed the code
                Else
                  pValue = String.Empty
                End If
              Catch vCareEX As CareException
                If vCareEX.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
                  ShowInformationMessage(vCareEX.Message)
                Else
                  Throw vCareEX
                End If
              Catch vEX As Exception
                Throw vEX
              End Try
            End If
          End If
          If pValue.Length > 0 Then CheckMailingCode(pValue, True)
        Case "SortOrder"
          Dim vEnabled As Boolean
          If eplStandard.FindComboBox("SortOrder").SelectedIndex <> SORT_BY_MAILSORT Then vEnabled = True
          cmdPrint.Enabled = vEnabled AndAlso mvMailingInfo.SummaryPrintValid
          cmdView.Enabled = vEnabled
        Case "CreateMailingHistory"
          eplStandard.FindTextLookupBox("Mailing").Text = ""
          eplStandard.FindTextBox("MailingCodeNotes").Text = ""
          eplStandard.FindTextBox("MailingCodeNotes").Visible = True
          eplStandard.PanelInfo.PanelItems("MailingCodeNotes").Hidden = False
          eplStandard.FindTextBox("MailingNotes").Text = ""
          eplStandard.FindDateTimePicker("Date").Value = CDate(AppValues.TodaysDate)
          eplStandard.FindTextLookupBox("Mailing").Enabled = Not eplStandard.FindCheckBox("CreateMailingHistory").Checked
          eplStandard.FindTextBox("MailingCodeNotes").Enabled = Not eplStandard.FindCheckBox("CreateMailingHistory").Checked
          eplStandard.FindTextBox("MailingNotes").Enabled = Not eplStandard.FindCheckBox("CreateMailingHistory").Checked
          eplStandard.FindDateTimePicker("Date").Enabled = Not eplStandard.FindCheckBox("CreateMailingHistory").Checked
          CheckMandatory()
        Case "StandardDocument"
          If eplStandard.FindTextLookupBox("StandardDocument").Text.ToString = "" Then
            eplStandard.FindCheckBox("AutoMerge").Enabled = False
            eplStandard.FindCheckBox("AutoMerge").Checked = False
          Else
            eplStandard.FindCheckBox("AutoMerge").Enabled = True
          End If
        Case "ExamCertRunType"
          CheckMandatory()
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub eplSOCancellation_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles eplSOCancellation.ValueChanged
    Try
      Select Case pParameterName
        Case "CancellationReason"
          CheckMandatory()
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub eplSelectionTester_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles eplSelectionTester.ValueChanged
    Try
      Select Case pParameterName
        Case "ReportDestination"
          CheckMandatory()
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub eplMisc_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles eplMisc.ValueChanged
    Try
      Dim vMemOrSub As Boolean
      vMemOrSub = (mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySubscriptions Or mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembers)

      Select Case pParameterName
        Case "Mailing"
          Dim vPos As Integer
          vPos = InStrRev(pValue, "|")
          If vPos > 0 Then
            pValue = Strings.Mid(pValue, vPos + 1)
          End If
          CheckMailingCode(pValue, True)
        Case "IncludeThankYou"
          eplMisc.FindCheckBox("DeleteThankYou").Checked = eplMisc.FindCheckBox("IncludeThankYou").Checked
          eplMisc.FindCheckBox("DeleteThankYou").Enabled = eplMisc.FindCheckBox("IncludeThankYou").Checked
          eplMisc.FindTextLookupBox("Company").Enabled = eplMisc.FindCheckBox("IncludeThankYou").Checked
          eplMisc.FindTextLookupBox("Mailing").Enabled = eplMisc.FindCheckBox("IncludeThankYou").Checked
          CheckMandatory()
        Case "CheckBoxOne"
          If vMemOrSub Then
            eplMisc.FindCheckBox("CheckBoxTwo").Enabled = Not eplMisc.FindCheckBox("CheckBoxOne").Checked
          End If
        Case "CheckBoxTwo"
          If vMemOrSub Then
            eplMisc.FindCheckBox("CheckBoxOne").Enabled = Not eplMisc.FindCheckBox("CheckBoxTwo").Checked
          End If
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub CheckMailingCode(ByVal pCode As String, ByVal pMailing As Boolean)
    'Given a mailing code check if it can be used
    'This is done by looking it up in the mailings table
    'and checking if the 'history_only' attribute is set
    'return True if the code is historic or False if not
    Dim vFound As Boolean
    If pCode.Length > 0 Then
      Dim vLookupList As New ParameterList(True)
      vLookupList("Mailing") = pCode
      Dim vRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMailings, vLookupList).Rows(0)

      If vRow IsNot Nothing Then
        vFound = True
        If pMailing Then
          If Convert.ToBoolean(CStr(IIf(vRow("HistoryOnly").ToString = "Y", True, False))) Then
            ShowInformationMessage(InformationMessages.ImMailingCodeHistoryOnly)
            eplStandard.FindTextBox("MailingCodeNotes").Enabled = False
          Else
            eplStandard.FindTextBox("MailingCodeNotes").Text = vRow("Notes").ToString
            eplStandard.FindTextBox("MailingCodeNotes").Enabled = False
          End If
        End If
      End If

      If Not vFound Then
        If pMailing Then
          eplStandard.FindTextLookupBox("Mailing").Text = ""
          If pCode.Length > 0 Then eplStandard.FindTextLookupBox("MailingCodeNotes").Focus()
        Else
          If pCode.Length > 0 Then ShowInformationMessage(InformationMessages.ImMailingCodeNotExists)
          eplMisc.FindTextLookupBox("Mailing").Text = ""
        End If
      End If
    End If
    CheckMandatory()
  End Sub

  Private Sub CheckMandatory()
    'Check that all the mandatory entries have been made on the form
    'Enable the OK button if so, otherwise disable it
    Dim vValid As Boolean = False
    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySelectionTester Then
      If eplSelectionTester.FindTextBox("ReportDestination").Text.Length > 0 Then vValid = True
    ElseIf mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandingOrderCancellation OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGAYECancellation Then
      If eplSOCancellation.FindTextLookupBox("CancellationReason").Text.Length > 0 Then vValid = True
    ElseIf mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyExamCertificates Then
      vValid = eplStandard.FindTextBox("ReportDestination").Text.Length > 0 And
               eplStandard.FindComboBox("ExamCertRunType").SelectedIndex > -1
    Else
      If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
        If eplStandard.FindTextLookupBox("Mailing").Text.Length > 0 Then 'Len(txtMailing(macMailing)) > 0 Then
          vValid = True
          If (eplStandard.FindComboBox("MailMerge").SelectedIndex <> 0) AndAlso eplStandard.FindTextLookupBox("Mailing").Text.Length = 0 Then
            vValid = False 'No destination
          End If
        ElseIf eplStandard.FindCheckBox("CreateMailingHistory").Checked = True AndAlso eplStandard.FindComboBox("MailMerge").SelectedIndex <> 0 AndAlso eplStandard.FindTextBox("ReportDestination").Text.Length > 0 Then
          vValid = True
        End If
      End If
      If vValid And mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then
        If eplMisc.FindCheckBox("IncludeThankYou").Checked = True Then ' chk(cacIncludeTYL).Value = vbChecked Then
          vValid = eplMisc.FindTextLookupBox("Company").Text.Length > 0 AndAlso eplMisc.FindTextLookupBox("Mailing").Text.Length > 0
        End If
      End If
    End If
    cmdOK.Enabled = vValid
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try

      Dim vResult As DialogResult = System.Windows.Forms.DialogResult.Yes
      Dim vNoOutput As Boolean
      Dim vDevice As String = ""
      Dim vParams As New ParameterList(True)
      Dim vSpecial As String = ""
      Dim vCombine As String = ""
      If mvOmitted IsNot Nothing Then
        CommitOmits()
      End If


      If eplStandard.Visible Then If eplStandard.AddValuesToList(vParams, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly) = False Then Exit Sub
      If eplMisc.Visible Then If eplMisc.AddValuesToList(vParams, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly) = False Then Exit Sub
      If eplSelectionTester.Visible Then If eplSelectionTester.AddValuesToList(vParams, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly) = False Then Exit Sub
      vParams = New ParameterList(True)

      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandingOrderCancellation Then
        CancelSOs()
      ElseIf mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGAYECancellation Then
        CancelGayePledges()
      Else
        'If the chosen sort order is Mailsort and the database is Oracle then ensure that the user's ODBC driver is configured to use the Disable Microsoft Transaction Server workaround
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_auto_name_mailing_files) Then
          If System.IO.File.Exists(eplStandard.FindTextBox("ReportDestination").Text) OrElse InStr(1, eplStandard.FindTextBox("ReportDestination").Text, "Mailing") = 0 Then
            eplStandard.FindTextBox("ReportDestination").Text = AppValues.GetMailingFileName(0)
          End If
          If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
            eplStandard.FindTextBox("ReportDestination").Text = eplStandard.FindTextBox("ReportDestination").Text.Replace("Mailing_", eplStandard.FindTextLookupBox("Mailing").Text & "_")
          Else
            eplStandard.FindTextBox("ReportDestination").Text = eplStandard.FindTextBox("ReportDestination").Text.Replace("Mailing_", CStr(eplStandard.FindComboBox("ExamCertRunType").SelectedValue) & "_")
          End If
        End If

        Dim vTextBox As TextBox = CType(IIf(eplSelectionTester.Visible = True, eplSelectionTester.FindTextBox("ReportDestination"), eplStandard.FindTextBox("ReportDestination")), TextBox)
        Dim vMailsort As Boolean = False
        Dim vSortOrder As ComboBox = eplStandard.FindPanelControl(Of ComboBox)("SortOrder", False)
        If vSortOrder IsNot Nothing AndAlso vSortOrder.SelectedIndex = SORT_BY_MAILSORT Then
          vMailsort = True
        End If

        If (eplStandard.FindComboBox("MailMerge").SelectedIndex > 0) And Not vMailsort Then
          'Don't do the following for mailsort because:
          '1/ CheckFilename opens the specified file.  So if it doesn't exist this will create it.
          '2/ the Application Parameters form that's displayed for mailsorted mailings checks for the file's existance as well.
          If CheckFileName(vTextBox.Text) = False Then
            vTextBox.Focus()
            Exit Sub
          End If
        End If

        If eplStandard.FindComboBox("MailMerge").SelectedIndex = 0 Then
          vResult = ShowQuestion(QuestionMessages.QmUpdateMailingHistoryNoOutput, MessageBoxButtons.YesNo)
          If vResult = System.Windows.Forms.DialogResult.No Then Exit Sub
          vNoOutput = True
        End If

        If vNoOutput = False Then
          vDevice = ""
          If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates AndAlso
             eplStandard.FindTextLookupBox("Device").Text.Length > 0 Then
            vDevice = eplStandard.FindTextLookupBox("Device").Text
          End If
        End If

        If vNoOutput Then
          vResult = System.Windows.Forms.DialogResult.Yes
        Else
          If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then
            ' Of the members that have been selected, output will only be produced for valid memberships that require a card.  Would you like to continue with the {0}?
            vResult = ShowQuestion(String.Format(QuestionMessages.QmMembershipCardsMemberSelected, "mailing"), MessageBoxButtons.YesNo)
          ElseIf mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyIrishGiftAid Then
            'vResult = AskQuestion(gmqIrishGiftAidAppropriateCertificates)
            ' Of the contacts that have been selected, output will only be produced for those contacts that have donated at least € {0} . Would you like to continue with the mailing?
            vResult = ShowQuestion(String.Format(QuestionMessages.QmIrishGAContactSelected, AppValues.ControlValue(AppValues.ControlValues.minimum_annual_donation)), MessageBoxButtons.YesNo)
          End If
          If vResult <> System.Windows.Forms.DialogResult.No Then
            Select Case mvMailingInfo.MailingType
              Case CareNetServices.MailingTypes.mtyMembers
                vCombine = CStr(IIf(eplMisc.FindCheckBox("CheckBoxOne").Checked, "Y", "N"))
                vSpecial = CStr(IIf(eplMisc.FindCheckBox("CheckBoxTwo").Checked, "Y", "N"))
              Case CareNetServices.MailingTypes.mtyMembershipCards, CareNetServices.MailingTypes.mtyPayers, CareNetServices.MailingTypes.mtyMemberFulfilment
                vSpecial = CStr(IIf(eplMisc.FindCheckBox("CheckBoxOne").Checked, "Y", "N"))
              Case CareNetServices.MailingTypes.mtySubscriptions
                vCombine = CStr(IIf(eplMisc.FindCheckBox("CheckBoxOne").Checked, "Y", "N"))
                vSpecial = CStr(IIf(eplMisc.FindCheckBox("CheckBoxTwo").Checked, "Y", "N"))
            End Select
          End If
        End If

        If vResult <> System.Windows.Forms.DialogResult.No Then
          vParams("SelectionSetNumber") = mvSelectionSet.ToString
          vParams("Revision") = mvMailingInfo.Revision.ToString
          vParams("ApplicationName") = CStr(IIf(mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards, "MC", mvMailingInfo.MailingTypeCode.ToString))
          vParams("RunPhase") = CStr(IIf(vNoOutput, "Phase3", "Phase2And3"))
          If mvSingleMembershipCardMailing Then vParams("RunPhase") = "All"
          vParams("Checkbox4") = CStr(IIf(eplMisc.FindCheckBox("CheckBoxOne").Checked, "Y", "N"))
          vParams("Checkbox5") = CStr(IIf(eplMisc.FindCheckBox("CheckBoxTwo").Checked, "Y", "N"))
          vParams("Special") = CStr(IIf(vSpecial.Length > 0, vSpecial, "N"))
          vParams("CombineMail") = CStr(IIf(vCombine.Length > 0, vCombine, "N"))
          vParams("Device") = vDevice
          vParams("DeviceDesc") = vDevice
          vParams.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
          vParams("ReportDestination") = vTextBox.Text
          If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
            vParams("SortOrder") = eplStandard.FindComboBox("SortOrder").Text
          End If
          vParams("Checkbox") = CStr(IIf(vMailsort, "Y", "N"))
          If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
            vParams("Mailing") = eplStandard.FindTextLookupBox("Mailing").Text
            If eplStandard.FindTextBox("MailingCodeNotes").Visible Then vParams("MailingDesc") = eplStandard.FindTextBox("MailingCodeNotes").Text
            If eplStandard.FindTextBox("MailingNotes").Visible Then vParams("Notes") = eplStandard.FindTextBox("MailingNotes").Text
            vParams("MailingDate") = eplStandard.FindDateTimePicker("Date").Value.ToString(AppValues.DateFormat)
          End If

          If eplSelectionTester.Visible = True Then
            vParams("Performance") = eplSelectionTester.FindTextLookupBox("Performance").Text
            vParams("Score") = eplSelectionTester.FindTextLookupBox("Score").Text
          End If
          vParams("RecordCount") = mvMailingInfo.SelectionCount.ToString()
          vParams("OrgLabelName") = mvMailingInfo.OrganisationLabelName

          If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
            vParams("NoMailingHistory") = CStr(IIf((eplStandard.FindCheckBox("CreateMailingHistory").Visible AndAlso eplStandard.FindCheckBox("CreateMailingHistory").Checked), "Y", "N"))
          Else
            vParams("NoMailingHistory") = "Y"
          End If

          If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards AndAlso mvIncludeTYLs Then
            vParams("TYLCompany") = ""
            If eplMisc.FindTextLookupBox("Company").Text.Length > 0 Then vParams("TYLCompany") = eplMisc.FindTextLookupBox("Company").Text
            vParams("TYLMailingCodes") = eplMisc.FindTextLookupBox("Mailing").Text ' , cftCharacter, BuildInValues(txtMailing(macTYLMailing).Text, "|"))
            vParams("TYLInclude") = CBoolYN(eplMisc.FindCheckBox("IncludeThankYou").Checked)
            vParams("TYLDelete") = CBoolYN(eplMisc.FindCheckBox("DeleteThankYou").Checked)

          End If
          vParams("UseStandardExclusions") = "N"
          vParams("BypassCount") = "N"
          vParams("GeneralMailing") = "Y"
          vParams("ScheduleMailing") = "Y"

          If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
            If eplStandard.FindTextLookupBox("StandardDocument").Visible AndAlso eplStandard.FindTextLookupBox("StandardDocument").ToString.Length > 0 Then
              vParams("StandardDocument") = eplStandard.FindTextLookupBox("StandardDocument").Text.ToString
              vParams("AutomaticMerge") = CBoolYN(eplStandard.FindCheckBox("AutoMerge").Checked)
              vParams("UserMailMerge") = "Y"
            End If
          Else
            Dim vExamSelector As ExamSelector = DirectCast(FindControl(eplStandard, "ExamUnitLinkId"), ExamSelector)
            vParams.IntegerValue("ExamUnitLinkId") = DirectCast(vExamSelector.SelectedNode.Tag, ExamSelectorItem).LinkID
            vParams("ExamCertRunType") = CStr(eplStandard.FindComboBox("ExamCertRunType").SelectedValue)
            vParams("AutomaticMerge") = CBoolYN(eplStandard.FindCheckBox("AutoMerge").Checked)
            vParams("SortOrder") = "ExamBookingUnitId"
            vParams("UserMailMerge") = "Y"
          End If

          Dim vContinue As Boolean = True
          If vMailsort Then
            'Get Mailsort parameters for mailsort report
            Dim vForm As New frmApplicationParameters(CareNetServices.TaskJobTypes.tjtMailingRun, vParams)
            vForm.ShowDialog()
            If vForm.DialogResult <> System.Windows.Forms.DialogResult.Cancel Then
              vParams.FillFromValueList(vForm.ReturnList.ValueList())
            Else
              vContinue = False
            End If
          End If

          If vContinue Then
            Dim vRunStatus As FormHelper.RunMailingResult = FormHelper.RunMailing(mvMailingInfo.TaskType, vParams)
            If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyExamCertificates AndAlso
               vRunStatus = FormHelper.RunMailingResult.MailingRunSynchSuccess AndAlso
               vParams("AutomaticMerge") = "Y" Then
              Call New ExamCertificateMergeEngine(eplStandard.FindPanelControl(Of TextBox)("ReportDestination").Text).ProduceDocuments()
            End If
          End If
          If mvSingleMembershipCardMailing Then Me.Close()
        End If
      End If
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub


  Public Function CheckFileName(ByVal pFileName As String) As Boolean
    Try
      'Returns True if filename is OK
      Dim vExit As Boolean

      If Strings.InStr(pFileName, "*") > 0 Or Strings.InStr(pFileName, "?") > 0 Then
        ShowInformationMessage(String.Format(InformationMessages.ImUnableToCreateOutputFile, pFileName))
        vExit = True
      End If
      Return Not vExit
    Catch
      ShowInformationMessage(String.Format(InformationMessages.ImUnableToCreateOutputFile, pFileName))
      Return False
    End Try
  End Function

  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
    Try
      mvMailingInfo.GenerateStatus = MailingInfo.MailingGenerateResult.mgrReset
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdRefine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefine.Click
    Try
      mvMailingInfo.GenerateStatus = MailingInfo.MailingGenerateResult.mgrRefine
      mvUnloadOK = True
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub frmGenMGen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    DisplayBlankGrid()
    SplitContainer1.Panel1.AutoScroll = True
    pnlEpl.AutoScroll = True

    If eplMisc.Visible Then
      eplMisc.Height = eplMisc.RequiredHeight
      eplStandard.Height = eplStandard.RequiredHeight
      eplStandard.Dock = DockStyle.Top
      eplMisc.Dock = DockStyle.Bottom
    ElseIf eplStandard.Visible Then
      eplStandard.Dock = DockStyle.Fill
    End If
  End Sub

  Private Sub cmdFindAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindAddress.Click
    Try
      Dim vFrmGenAddress As New frmGenMAddress(mvContactNumber, mvAddressNumber, mvMailingInfo)
      vFrmGenAddress.ShowDialog()

      If vFrmGenAddress.AddressNumber > 0 Then
        ' Replace address number of selected row of datagrid
        If drg.CurrentDataRow > 0 Then
          drg.SetValue(drg.CurrentDataRow, "address_number", vFrmGenAddress.AddressNumber.ToString)
          SelectViewListItem(drg.CurrentDataRow)
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
    Try
      Dim vParams As New ParameterList(True)
      SetStatus(ControlText.FrmGeneralMailingSelectContacts)    'Generating Selected Contacts
      vParams("SelectionSetNumber") = mvSelectionSet.ToString
      DataHelper.DeleteSelectedContact(vParams)
      DataHelper.AddMailingSelectedContacts(mvMailingInfo.SelectionSet, mvMailingInfo.Revision, mvMailingInfo.SelectionTable)
      ClearStatus()
      Dim vRDS As New frmReportDataSelection(mvSelectionSet, False)
      vRDS.ShowDialog()
      SetStatus(ControlText.FrmGeneralMailingRemoveContacts)    'Removing Selected Contacts
      DataHelper.DeleteSelectedContact(vParams)
      ClearStatus()
      Exit Sub

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdMerge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMerge.Click
    Try
      Dim vCaptions() As String
      Dim vCustomData As String = ""
      Dim vCustom As Boolean
      Dim vAttrCaptions As String
      Dim vIndex As Integer
      Dim vResult As DialogResult
      Dim vRecords As Integer

      Dim vfrmGenList As New frmGenMLists(mvMailingTypeCode, mvMailingInfo)
      vfrmGenList.ShowDialog()

      If mvMailingInfo.NewSelectionSet > 0 Then
        vResult = ShowQuestion(QuestionMessages.QmDeDuplicationContactAddress, MessageBoxButtons.YesNoCancel)
        If vResult = System.Windows.Forms.DialogResult.No Or vResult = System.Windows.Forms.DialogResult.Yes Then
          Dim vParams As New ParameterList(True)
          vParams("SelectionSetNumber") = mvMailingInfo.NewSelectionSet.ToString
          Dim vDataSet As DataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstSelectionSetData, CareServices.XMLTableDataSelectionTypes), vParams)
          If vDataSet IsNot Nothing Then
            vCustom = CBool(IIf(vDataSet.Tables("DataRow").Rows(0).Item("CustomData").ToString = "", False, vDataSet.Tables("DataRow").Rows(0).Item("CustomData")))
            vAttrCaptions = CStr(vDataSet.Tables("DataRow").Rows(0).Item("AttributeCaptions"))
            If vCustom Then
              vCaptions = Split(vAttrCaptions, ",")
              For vIndex = 1 To UBound(vCaptions)
                vCustomData = vCustomData & ","
              Next
            End If
          End If

          vParams("UserName") = DataHelper.UserInfo.Logname 'pList("Owner").ToString
          vParams("Department") = DataHelper.UserInfo.Department  'pList("Department").ToString
          vParams("SelectionSetDesc") = vDataSet.Tables("DataRow").Rows(0).Item("SelectionSetDesc").ToString
          vParams.IntegerValue("NumberInMailing") = mvMailingInfo.SelectionCount
          vParams("ApplicationName") = mvMailingInfo.MailingTypeCode
          vParams.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
          vParams.IntegerValue("Revision") = 1
          vParams.IntegerValue("NewSelectionSetNumber") = mvMailingInfo.NewSelectionSet
          vParams("Unique") = CStr(IIf(vResult = System.Windows.Forms.DialogResult.Yes, "Y", "N"))
          DataHelper.CreateCopyOfSelectionSet(vParams)

          If vCustom Then
            Dim vDeDuplication As String = CStr(IIf(vResult = System.Windows.Forms.DialogResult.Yes, "Y", "N"))
            DataHelper.AddSelectionSetData(mvMailingInfo.NewSelectionSet, vCustomData, vDeDuplication)
          End If
          vParams = New ParameterList(True)
          vParams.IntegerValue("SelectionSetNumber") = mvMailingInfo.NewSelectionSet
          vRecords = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctSelectedContacts, vParams)

          Dim vSelectionSetParams As New ParameterList(True)
          vSelectionSetParams("SelectionSetNumber") = mvMailingInfo.NewSelectionSet.ToString
          vSelectionSetParams("NumberInSet") = CStr(vRecords)
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vSelectionSetParams)
          ShowInformationMessage(InformationMessages.ImMergeCompleted, CStr(vRecords))
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSaveList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveList.Click
    Try
      ProcessSave(SaveTypes.List)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub ProcessSave(ByVal pSaveType As SaveTypes)
    Dim vParams As New ParameterList(True)
    Dim vCaption As String = "Save"
    Dim vSaveResults As ParameterList
    vParams("ApplicationName") = mvMailingTypeCode
    Select Case pSaveType
      Case SaveTypes.List
        vParams("SaveType") = "List"
        vCaption = ControlText.FrmGenMSaveList
      Case SaveTypes.CriteriaSet
        vParams("SaveType") = "Criteria"
        vCaption = ControlText.FrmGenMSaveCriteria
      Case Else
        vParams("SaveType") = ""
    End Select
    vSaveResults = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optGenMSave, Nothing, vParams, vCaption)
    If vSaveResults IsNot Nothing Then ProcessSave(vSaveResults, pSaveType)
  End Sub

  Private Sub ProcessSave(ByVal pList As ParameterList, ByVal pSaveType As SaveTypes)
    If pList("CriteriaSetDesc").ToString.Length > 0 Then
      Select Case pSaveType
        Case SaveTypes.List
          Dim vParams As New ParameterList(True)
          vParams("UserName") = DataHelper.UserInfo.Logname 'pList("Owner").ToString
          vParams("Department") = DataHelper.UserInfo.Department  'pList("Department").ToString
          vParams("SelectionSetDesc") = pList("CriteriaSetDesc").ToString
          vParams.IntegerValue("NumberInMailing") = mvMailingInfo.SelectionCount
          vParams("ApplicationName") = mvMailingInfo.MailingTypeCode
          vParams.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
          vParams("SaveSelectionSet") = "Y"

          Dim vResult As ParameterList = DataHelper.CreateCopyOfSelectionSet(vParams)

          If Not vResult.Contains("Result") OrElse vResult("Result") <> "OK" Then
            ShowInformationMessage(InformationMessages.ImProblemSavingSelectionSet)
          End If
        Case SaveTypes.CriteriaSet
          Dim vInsertParams As New ParameterList(True)
          vInsertParams.IntegerValue("CriteriaSetNumber") = 0
          vInsertParams("CriteriaSetDesc") = pList("CriteriaSetDesc")
          vInsertParams("UserName") = pList("UserName")
          vInsertParams("Department") = pList("Department")
          If pList.Contains("ReportCode") AndAlso pList("ReportCode").Length > 0 Then vInsertParams("ReportCode") = pList("ReportCode")
          If pList.Contains("StandardDocument") AndAlso pList("StandardDocument").Length > 0 Then vInsertParams("StandardDocument") = pList("StandardDocument")
          vInsertParams("ApplicationName") = mvMailingInfo.MailingTypeCode
          Dim vReturnList As ParameterList = DataHelper.AddCriteriaSet(vInsertParams)

          Dim vList As New ParameterList(True)
          vList.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
          Dim vdsTemp As DataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstCriteriaSetDetails, CareServices.XMLTableDataSelectionTypes), vList)
          If vdsTemp IsNot Nothing AndAlso vdsTemp.Tables.Contains("DataRow") Then
            For Each vDataRow As DataRow In vdsTemp.Tables("DataRow").Rows
              vInsertParams = New ParameterList(True)
              vInsertParams.IntegerValue("CriteriaSet") = vReturnList.IntegerValue("CriteriaSetNumber")
              vInsertParams.IntegerValue("SequenceNumber") = IntegerValue(vDataRow("SequenceNumber"))
              vInsertParams("AndOr") = vDataRow("AndOr").ToString
              vInsertParams("LeftParenthesis") = vDataRow("LeftParenthesis").ToString
              vInsertParams("IE") = vDataRow("IE").ToString
              vInsertParams("CO") = vDataRow("CO").ToString
              vInsertParams("SearchArea") = vDataRow("SearchArea").ToString
              vInsertParams("MainValue") = vDataRow("MainValue").ToString
              vInsertParams("SubsidiaryValue") = vDataRow("SubsidiaryValue").ToString
              vInsertParams("Period") = vDataRow("Period").ToString
              vInsertParams("RightParenthesis") = vDataRow("RightParenthesis").ToString
              DataHelper.AddCriteriaSetDetails(vInsertParams)
            Next
          End If
      End Select
    End If
  End Sub

  Private Sub cmdSaveCriteria_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveCriteria.Click
    Try
      ProcessSave(SaveTypes.CriteriaSet)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
    Dim vCount As Integer
    Dim vResult As DialogResult

    Try
      vCount = mvMailingInfo.GetMailingSelectionCount(mvMailingInfo.SelectionSet, mvMailingInfo.Revision, mvMailingTypeCode) 'mvMailingSelection.Revision)
      If vCount > 1000 Then
        vResult = ShowQuestion(QuestionMessages.QmViewCount, MessageBoxButtons.OKCancel, vCount.ToString)
      Else
        vResult = System.Windows.Forms.DialogResult.OK
      End If
      If vResult = System.Windows.Forms.DialogResult.OK Then
        Select Case mvMailingInfo.MailingType
          Case CareNetServices.MailingTypes.mtyMembers, CareNetServices.MailingTypes.mtyMembershipCards
            If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then ShowInformationMessage(InformationMessages.ImMembershipCardView)
            GetRelevantMembers()
          Case CareNetServices.MailingTypes.mtyExamCertificates
            GetRelevantExamBookings()
          Case Else
            GetSelectedContacts()
        End Select
      End If
      Exit Sub

    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

  Private Sub GetRelevantMembers(Optional ByVal pView As Boolean = True)
    'This method is copied to the ScheduledMailing Class for scheduled mailings, any changes here should be replicated in there as well
    Dim vParams As New ParameterList(True)
    vParams.IntegerValue("SelectionSetNumber") = mvSelectionSet
    vParams.IntegerValue("Revision") = mvMailingInfo.Revision
    vParams("ApplicationCode") = mvMailingTypeCode
    vParams("Address") = mvMailingInfo.OrganisationMailTo
    vParams("AddressUsage") = mvMailingInfo.OrganisationAddressUsage
    vParams("Checkbox") = CStr(IIf(eplMisc.FindCheckBox("CheckBoxOne").Checked, "Y", "N")) '"Y"
    vParams("Checkbox2") = CStr(IIf(eplMisc.FindCheckBox("CheckBoxTwo").Checked, "Y", "N"))  '"Y"
    vParams("OrderBy") = eplStandard.FindComboBox("SortOrder").Text

    Dim vDataSet As DataSet = DataHelper.GetMailingRelevantMembers(vParams)
    If vDataSet IsNot Nothing Then

      If Not vDataSet.Tables.Contains("Column") Then
        Dim vTable As DataTable = DataHelper.NewColumnTable

        vDataSet.Tables.Add(vTable)

        If mvMailingInfo.MasterAttribute <> "contact_number" Then
          DataHelper.AddDataColumn(vTable, mvMailingInfo.MasterAttribute, "Master", , "N")
        End If
        DataHelper.AddDataColumn(vTable, "contact_number", "Number")
        DataHelper.AddDataColumn(vTable, "address_number", "Address No")
        DataHelper.AddDataColumn(vTable, "label_name", "Name")
        DataHelper.AddDataColumn(vTable, "department", "Dept")
      End If

      SetDepartmentInGrid(vDataSet)
      drg.Populate(vDataSet)
    End If

    SelectViewListItem(0)
  End Sub

  Private Sub GetRelevantExamBookings(Optional ByVal pView As Boolean = True)
    'This method is copied to the ScheduledMailing Class for scheduled mailings, any changes here should be replicated in there as well
    Dim vParams As New ParameterList(True)
    vParams.IntegerValue("SelectionSetNumber") = mvSelectionSet
    vParams.IntegerValue("Revision") = mvMailingInfo.Revision
    vParams("ApplicationCode") = mvMailingTypeCode
    vParams.IntegerValue("ExamUnitLinkId") = DirectCast(DirectCast(FindControl(eplStandard, "ExamUnitLinkId"), ExamSelector).SelectedNode.Tag, ExamSelectorItem).LinkID
    vParams("ExamCertRunType") = CStr(eplStandard.FindComboBox("ExamCertRunType").SelectedValue)

    Dim vDataSet As DataSet = DataHelper.GetMailingRelevantExamBookingUnits(vParams)
    If vDataSet IsNot Nothing Then

      If Not vDataSet.Tables.Contains("Column") Then
        Dim vTable As DataTable = DataHelper.NewColumnTable

        vDataSet.Tables.Add(vTable)

        DataHelper.AddDataColumn(vTable, "exam_student_unit_header_id", "Student Unit Header Id")
        DataHelper.AddDataColumn(vTable, "contact_number", "Number")
        DataHelper.AddDataColumn(vTable, "address_number", "Address No")
        DataHelper.AddDataColumn(vTable, "label_name", "Name")
      End If

      drg.Populate(vDataSet)
    End If

    SelectViewListItem(0)
  End Sub

  Private Sub drg_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles drg.RowSelected
    Try
      SelectViewListItem(pRow)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub GetSelectedContacts()
    Dim vParams As New ParameterList(True)
    vParams.IntegerValue("SelectionSetNumber") = mvSelectionSet
    vParams.IntegerValue("Revision") = mvMailingInfo.Revision
    vParams("ApplicationCode") = mvMailingTypeCode
    vParams("Address") = mvMailingInfo.OrganisationMailTo
    vParams("AddressUsage") = mvMailingInfo.OrganisationAddressUsage
    vParams("Checkbox") = CStr(IIf(eplMisc.FindCheckBox("CheckBoxOne").Checked, "Y", "N")) '"Y"
    vParams("Checkbox2") = CStr(IIf(eplMisc.FindCheckBox("CheckBoxTwo").Checked, "Y", "N"))  '"Y"
    vParams("OrderBy") = eplStandard.FindComboBox("SortOrder").Text

    Dim vDataSet As DataSet = DataHelper.GetMailingSelectedContacts(vParams)
    SetDepartmentInGrid(vDataSet)
    If vDataSet IsNot Nothing Then

      If Not vDataSet.Tables.Contains("Column") Then
        Dim vTable As DataTable = DataHelper.NewColumnTable

        vDataSet.Tables.Add(vTable)
        If mvMailingInfo.MasterAttribute <> "contact_number" Then
          DataHelper.AddDataColumn(vTable, mvMailingInfo.MasterAttribute, "Master", , "N")
        End If
        DataHelper.AddDataColumn(vTable, "contact_number", "Number")
        DataHelper.AddDataColumn(vTable, "address_number", "Address No")
        DataHelper.AddDataColumn(vTable, "label_name", "Name")
        DataHelper.AddDataColumn(vTable, "department", "Dept")
      End If

      SetDepartmentInGrid(vDataSet)
      drg.Populate(vDataSet)
    End If

    SelectViewListItem(0)
  End Sub

  Private Sub SetDepartmentInGrid(ByVal pDataSet As DataSet)
    If pDataSet IsNot Nothing Then
      If pDataSet.Tables.Contains("DataRow") Then
        For Each vDataRow As DataRow In pDataSet.Tables("DataRow").Rows
          If vDataRow("Department").ToString() = DataHelper.UserInfo.Department Then
            vDataRow("Department") = String.Empty
          End If
        Next
      End If
    End If
  End Sub

  Private Sub DisplayBlankGrid()
    If pnlGrid.Visible = True Then
      Dim vDataSet As DataSet = New DataSet()
      Dim vTable As DataTable = DataHelper.NewColumnTable

      vDataSet.Tables.Add(vTable)
      If Not mvMailingInfo.MasterAttribute.Equals("contact_number", StringComparison.InvariantCultureIgnoreCase) AndAlso
         Not mvMailingInfo.MasterAttribute.Equals("exam_student_unit_header_id", StringComparison.InvariantCultureIgnoreCase) Then
        DataHelper.AddDataColumn(vTable, mvMailingInfo.MasterAttribute, "Master", , "N")
      End If
      If mvMailingInfo.MasterAttribute.Equals("exam_student_unit_header_id", StringComparison.InvariantCultureIgnoreCase) Then
        DataHelper.AddDataColumn(vTable, "exam_student_unit_header_id", "Student Unit Header Id")
        DataHelper.AddDataColumn(vTable, "contact_number", "Number")
        DataHelper.AddDataColumn(vTable, "address_number", "Address No")
        DataHelper.AddDataColumn(vTable, "label_name", "Name")
      Else
        DataHelper.AddDataColumn(vTable, "contact_number", "Number")
        DataHelper.AddDataColumn(vTable, "address_number", "Address No")
        DataHelper.AddDataColumn(vTable, "label_name", "Name")
        DataHelper.AddDataColumn(vTable, "department", "Dept")
      End If

      'define the columns for the datarow table
      AddDataRowTableToDataSet(vDataSet)

      drg.Populate(vDataSet)
    End If
  End Sub

  Private Sub SelectViewListItem(ByVal pRow As Integer)
    Dim vContactNumber As Integer
    Dim vAddressNumber As Integer

    With drg
      If pRow > -1 AndAlso .RowCount > 0 Then
        vContactNumber = IntegerValue(.GetValue(pRow, "contact_number"))
        vAddressNumber = IntegerValue(.GetValue(pRow, "address_number"))
        If vContactNumber > 0 Then
          Dim vContactInfo As ContactInfo = New ContactInfo(vContactNumber)
          If vContactInfo.ContactType = ContactInfo.ContactTypes.ctOrganisation Then
            Dim vList As ParameterList = New ParameterList(True)
            vList.IntegerValue("ContactNumber") = vContactNumber
            vList.IntegerValue("AddressNumber") = vAddressNumber
            Dim vDataset As DataSet = DataHelper.GetAddressData(CareServices.XMLAddressDataSelectionTypes.xadtAddressInformation, vList)
            If vDataset IsNot Nothing AndAlso vDataset.Tables.Contains("DataRow") Then
              lblAddress.Text = vDataset.Tables("DataRow").Rows(0)("AddressLine").ToString
            End If
            lblOrganisation.Text = vContactInfo.ContactName
          Else
            Dim vList As New ParameterList(True)
            vList.IntegerValue("ContactNumber") = vContactNumber
            vList.IntegerValue("AddressNumber") = vAddressNumber
            Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressInformation, vContactNumber, vList)
            If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then
              lblAddress.Text = vDataSet.Tables("DataRow").Rows(0)("AddressLine").ToString
              If IntegerValue(vDataSet.Tables("DataRow").Rows(0)("OrganisationNumber")) > 0 Then
                Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressPositionAndOrganisation, vContactNumber, vList))
                If vRow IsNot Nothing Then lblOrganisation.Text = vRow.Item("Name").ToString
              Else
                lblOrganisation.Text = ""
              End If
            End If
          End If
          cmdOmit.Enabled = True
          cmdFindAddress.Enabled = True
          mvContactNumber = vContactNumber
          mvAddressNumber = vAddressNumber
          mvCaption = mvMailingInfo.Caption
        Else
          lblAddress.Text = ""
          lblOrganisation.Text = ""
          cmdOmit.Enabled = False
          cmdFindAddress.Enabled = False
        End If
      End If
    End With
  End Sub

  Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
    Dim vContinue As Boolean
    Dim vResult As ParameterList = Nothing
    Dim vSortFields As String
    Try
      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembers Then
        vResult = ShowMemPrintForm(CareNetServices.MailingTypes.mtyMembers)
      ElseIf mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then
        If ShowQuestion(QuestionMessages.QmMembershipCardsMemberSelected, MessageBoxButtons.YesNo, QuestionMessages.QmSummaryReport) = System.Windows.Forms.DialogResult.Yes Then    'summary report
          vResult = ShowMemPrintForm(CareNetServices.MailingTypes.mtyMembershipCards)
        End If
      Else
        vContinue = True
      End If
      If vContinue OrElse vResult IsNot Nothing Then
        If vResult IsNot Nothing Then mvReportOption = CType(vResult.IntegerValue("ReportOption"), ReportOptions)
        vSortFields = eplStandard.FindComboBox("SortOrder").Text
        GMPrintSummary(vSortFields, eplStandard.FindTextBox("SummaryTitle").Text)
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub CancelSOs()
    Dim vParms As New ParameterList(True)
    vParms.IntegerValue("SelectionSetNumber") = mvSelectionSet
    vParms.IntegerValue("Revision") = mvMailingInfo.Revision
    vParms("ApplicationCode") = mvMailingInfo.MailingTypeCode
    vParms("TableName") = mvMailingInfo.SelectionTable
    vParms("CancellationReason") = eplSOCancellation.FindTextLookupBox("CancellationReason").Text
    vParms("OrderBy") = ""
    DataHelper.CancelStandingOrders(vParms)
    cmdPrint_Click(Me, Nothing)
    ShowInformationMessage(InformationMessages.ImSOCancellationComplete) ' Standing Order Cancellation Complete
  End Sub

  Private Sub ClearStatus()
    Me.Text = mvMailingInfo.Caption + ControlText.FrmGeneralMailing
  End Sub

  Private Sub SetStatus(ByVal pText As String)
    Me.Text = pText
  End Sub

  Private Sub cmdOmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOmit.Click
    Try
      If drg.RowCount > 0 Then
        Dim vdrSelection As DataRow = drg.DataSourceDataRow(drg.CurrentDataRow)
        Dim vdtSelection As DataTable = vdrSelection.Table
        Dim vNumber As Integer = CInt(vdrSelection(mvMailingInfo.MasterAttribute))
        If vNumber > 0 Then
          If mvOmitted Is Nothing Then
            mvOmitted = New List(Of Integer)
          End If
          mvOmitted.Add(vNumber)
          vdrSelection.Delete()
          vdrSelection.Table.AcceptChanges()
          mvMailingInfo.SelectionCount = mvMailingInfo.SelectionCount - 1
          If mvMailingInfo.SelectionCount < 1 Then
            cmdOK.Enabled = False
            cmdReport.Enabled = False
            cmdSaveList.Enabled = False
            cmdPrint.Enabled = False
            cmdView.Enabled = False
            SelectViewListItem(-1)
          Else
            SelectViewListItem(drg.CurrentDataRow)
          End If

          If drg.RowCount = 0 Then
            lblAddress.Text = ""
            lblOrganisation.Text = ""
            cmdOmit.Enabled = False
            cmdFindAddress.Enabled = False
          End If
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub CancelGayePledges()
    Dim vParms As New ParameterList(True)
    vParms.IntegerValue("SelectionSetNumber") = mvSelectionSet
    vParms.IntegerValue("Revision") = mvMailingInfo.Revision
    vParms("ApplicationCode") = mvMailingInfo.MailingTypeCode
    vParms("TableName") = mvMailingInfo.SelectionTable
    vParms("CancellationReason") = eplSOCancellation.FindTextLookupBox("CancellationReason").Text 'eplStandard.FindComboBox("SortOrder").Text
    DataHelper.CancelGayePledges(vParms)
    cmdPrint_Click(Me, Nothing)
    ShowInformationMessage(InformationMessages.ImPGPCancellationComplete) ' Payroll Giving Pledge Cancellation Complete
  End Sub

  Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click
    Try
      Dim vParameters As New Collection
      Dim vReportParams As New ParameterList(True)

      vReportParams("RP1") = mvMailingInfo.SelectionTable
      vReportParams.IntegerValue("RP2") = mvSelectionSet
      vReportParams.IntegerValue("RP3") = mvMailingInfo.Revision
      vReportParams("ReportCode") = "SMQC"
      Call (New PrintHandler).PrintReport(vReportParams, "")

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function ShowMemPrintForm(ByVal pGeneralMailing As CareNetServices.MailingTypes) As ParameterList
    Dim vPanelItems As New PanelItems("epl")
    Dim vDefault As New ParameterList(True)
    Dim vResult As ParameterList = Nothing
    Select Case pGeneralMailing
      Case CareNetServices.MailingTypes.mtyMembers
        vResult = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optGenMPrint, vPanelItems, vDefault, "Member Print")
      Case CareNetServices.MailingTypes.mtyMembershipCards
        vResult = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optGenMPrint, vPanelItems, vDefault, "Membership Card Print")
    End Select
    Return vResult
  End Function


  Public Sub GMPrintSummary(ByVal pSortFields As String, ByVal pTitle As String)
    Dim vType As String
    Dim vReportList As New ParameterList(True)

    vReportList("RP1") = mvMailingInfo.SelectionTable
    vReportList("RP2") = mvSelectionSet.ToString
    vReportList("RP3") = mvMailingInfo.Revision.ToString

    If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyStandingOrderCancellation And _
       mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyGAYECancellation Then
      vReportList("RP4") = pSortFields

      If pTitle.Length = 0 Then
        Select Case mvMailingInfo.MailingType
          Case CareNetServices.MailingTypes.mtyMembers, CareNetServices.MailingTypes.mtyMembershipCards
            If mvReportOption = ReportOptions.roDetailed Then
              pTitle = "Detailed Member Print"
            ElseIf mvReportOption = ReportOptions.roRegister Then
              pTitle = "Register of Members"
            Else  'roSummary
              pTitle = "Summary Member Print"
            End If
          Case Else
            pTitle = "Summary Print"
        End Select
      End If
      vReportList("RP5") = pTitle

      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtySubscriptions Then
        vReportList("RP6") = DataHelper.UserInfo.Department
      End If

      vType = "SP"
      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembers Or _
         mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then
        If mvReportOption = ReportOptions.roDetailed Then
          vType = "DP"
        ElseIf mvReportOption = ReportOptions.roRegister Then
          vType = "RP"
        ElseIf pSortFields.Length > 8 AndAlso pSortFields.Substring(0, 8) = "m.branch" Then
          vType = "BP"
        End If
      End If

      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then AddCardParameters(vReportList)
    Else
      vType = "CN"
    End If

    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards Then
      vType = "MC" & vType
    Else
      vType = mvMailingInfo.MailingTypeCode & vType
    End If

    If eplMisc.Visible Then
      If eplMisc.FindCheckBox("CheckBoxOne").Checked Then vReportList("Checkbox4") = "Y"
      If eplMisc.FindCheckBox("CheckBoxTwo").Checked Then vReportList("Checkbox5") = "Y"
    End If

    vReportList("ReportCode") = vType
    vReportList("ApplicationName") = mvMailingInfo.MailingTypeCode
    vReportList.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
    Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
  End Sub

  Private Sub AddCardParameters(ByRef pParams As ParameterList)
    Dim vTemp(3) As Boolean
    vTemp(mvMailingInfo.CardProductionType) = True
    pParams("RP10") = IIf(vTemp(MembershipCardProductionTypes.mcpDefault) = True, "Y", "").ToString
    pParams("RP11") = IIf(vTemp(MembershipCardProductionTypes.mcpAutoOrPaid) = True, "Y", "").ToString
    pParams("RP12") = IIf(vTemp(MembershipCardProductionTypes.mcpPaymentRequired) = True, "Y", "").ToString
    pParams("prefix_mem_type") = IIf(AppValues.ConfigurationOption(AppValues.ConfigurationOptions.prefix_member_with_type, True), "Y", "N").ToString
  End Sub

  Private Sub ShowMiscCheckBoxes(ByVal pVisible As Boolean)
    eplMisc.FindCheckBox("CheckBoxOne").Visible = pVisible
    eplMisc.PanelInfo.PanelItems("CheckBoxOne").Hidden = Not pVisible
    eplMisc.FindLabel("CheckBoxOne").Visible = pVisible
    eplMisc.FindCheckBox("CheckBoxTwo").Visible = pVisible
    eplMisc.PanelInfo.PanelItems("CheckBoxTwo").Hidden = Not pVisible
    eplMisc.FindLabel("CheckBoxTwo").Visible = pVisible
  End Sub

  Private Sub frmGenMGen_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    Try
      If mvUnloadOK = False Then
        'Cancel or close window
        mvMailingInfo.DeleteSelection(mvMailingInfo.SelectionSet, mvMailingInfo.Revision)
        If mvMailingInfo.Revision > 0 Then mvMailingInfo.Revision = mvMailingInfo.Revision - 1
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Try
      If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
        eplStandard.FindComboBox("SortOrder").SelectedIndex = DEF_SORT_ORDER
      End If
      eplStandard.FindComboBox("MailMerge").SelectedIndex = 1

      If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
        eplStandard.FindTextLookupBox("Mailing").Text = ""
        eplMisc.FindTextLookupBox("Mailing").Text = ""
      End If

      eplStandard.FindTextBox("ReportDestination").Text = ""
      eplSelectionTester.FindTextBox("ReportDestination").Text = ""
      If mvMailingInfo.MailingType <> CareNetServices.MailingTypes.mtyExamCertificates Then
        eplStandard.FindTextBox("MailingNotes").Text = ""
      End If
      If drg.RowCount > 0 Then
        drg.ClearDataRows()
      End If
      DisplayBlankGrid()
      lblOrganisation.Text = ""
      lblAddress.Text = ""
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub drg_LeaveCell(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer) Handles drg.RowSelected
    Try
      SelectViewListItem(pRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  ''' <summary>
  ''' Delete items Omitted. 
  ''' </summary>
  ''' <remarks>Passes a CSV to the server</remarks>
  Private Sub CommitOmits()

    Dim vStringbuilder As New StringBuilder

    For Each vNum As Integer In mvOmitted
      vStringbuilder.Append(vNum.ToString)
      vStringbuilder.Append(",")
    Next
    vStringbuilder.Remove(vStringbuilder.Length - 1, 1)
    DataHelper.DeleteMailingSelectionSet(mvSelectionSet, mvMailingInfo.MailingTypeCode, mvMailingInfo.Revision, False, vStringbuilder.ToString)
  End Sub

End Class
