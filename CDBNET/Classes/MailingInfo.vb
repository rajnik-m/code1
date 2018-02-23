Imports System.Text.RegularExpressions

Public Class MailingInfo
  Private mvCurrentCriteriaValid As Boolean
  Private mvMailingType As CareNetServices.MailingTypes
  Private mvCriteriaRows As Integer
  Private mvCriteriaDetails As CriteriaDetails
  Private mvRevision As Integer
  Private mvCriteriaSet As Integer
  Private mvNewCriteriaSet As Integer
  Private mvMailingTypeCode As String
  Private mvAppealMailing As Boolean
  Private mvCaption As String
  Private mvSelectionTable As String
  Private mvNewSelectionSet As Integer 'Selected from the Selection Sets List
  Private mvSelectionSet As Integer         'Current Selection Set
  Private mvBypassCriteriaCount As Boolean
  Private mvGenerateStatus As MailingGenerateResult
  Private mvCount As Integer
  Private mvSelectionCount As Integer
  Private mvMasterAttribute As String
  Private mvExclusionCriteriaSet As Integer
  Private mvReportType As ReportOptions
  Private mvVariableParameters As String = ""
  Private mvOrganisationMailTo As String = ""
  Private mvOrganisationMailWhere As String = ""
  Private mvOrganisationLabelName As String = ""
  Private mvOrganisationAddressUsage As String = ""
  Private mvOrganisationRoles As String = ""
  Private mvIncludeHistoricRoles As Boolean
  Private mvMailingSelection As Boolean
  Private mvOrganisationCriteriaCount As Integer
  Private mvDisplayOrgSelection As Boolean
  Private mvCardProductionType As MembershipCardProductionTypes
  Private mvSummaryPrintValid As Boolean
  Private mvTaskType As CareServices.TaskJobTypes

  Public Enum MailingGenerateResult
    mgrNone = 1
    mgrRefine
    mgrReset
  End Enum
  Public Enum ReportOptions
    roDetailed = 1
    roSummary
    roRegister
  End Enum

  Public Enum MembershipCardProductionTypes
    mcpDefault = 1
    mcpAutoOrPaid
    mcpPaymentRequired
  End Enum

  Public Property ExclusionCriteriaSet() As Integer
    Get
      Return mvExclusionCriteriaSet
    End Get
    Set(ByVal value As Integer)
      mvExclusionCriteriaSet = value
    End Set
  End Property

  Public Property Count() As Integer
    Get
      Return mvCount
    End Get
    Set(ByVal value As Integer)
      mvCount = value
    End Set
  End Property

  Public Property SelectionCount() As Integer
    Get
      Return mvSelectionCount
    End Get
    Set(ByVal value As Integer)
      mvSelectionCount = value
    End Set
  End Property

  Public Property GenerateStatus() As MailingGenerateResult
    Get
      Return mvGenerateStatus
    End Get
    Set(ByVal value As MailingGenerateResult)
      mvGenerateStatus = value
    End Set
  End Property

  Public Property BypassCriteriaCount() As Boolean
    Get
      Return mvBypassCriteriaCount
    End Get
    Set(ByVal value As Boolean)
      mvBypassCriteriaCount = value
    End Set
  End Property

  Public Property SelectionSet() As Integer
    Get
      Return mvSelectionSet
    End Get
    Set(ByVal value As Integer)
      mvSelectionSet = value
    End Set
  End Property

  Public Property NewSelectionSet() As Integer
    Get
      Return mvNewSelectionSet
    End Get
    Set(ByVal value As Integer)
      mvNewSelectionSet = value
    End Set
  End Property
  Public ReadOnly Property MasterAttribute() As String
    Get
      Return mvMasterAttribute
    End Get
  End Property

  Public ReadOnly Property Caption() As String
    Get
      Return mvCaption
    End Get
  End Property

  Public ReadOnly Property MailingTypeCode() As String
    Get
      Return mvMailingTypeCode
    End Get
  End Property

  Public Property CurrentCriteriaValid() As Boolean
    Get
      Return mvCurrentCriteriaValid
    End Get
    Set(ByVal value As Boolean)
      mvCurrentCriteriaValid = value
    End Set
  End Property

  Public ReadOnly Property AppealMailing() As Boolean
    Get
      Return mvAppealMailing
    End Get
  End Property

  Public Property CriteriaRows() As Integer
    Get
      Return mvCriteriaRows
    End Get
    Set(ByVal value As Integer)
      mvCriteriaRows = value
    End Set
  End Property

  Public ReadOnly Property MailingType() As CareNetServices.MailingTypes
    Get
      Return mvMailingType
    End Get
  End Property

  Public ReadOnly Property CurrentCriteria() As CriteriaDetails
    Get
      If mvCriteriaDetails Is Nothing Then
        mvCriteriaDetails = New CriteriaDetails
        'mvCriteriaDetails.Init()
      End If
      Return mvCriteriaDetails
    End Get
  End Property

  Public Property Revision() As Integer
    Get
      Return mvRevision
    End Get
    Set(ByVal value As Integer)
      mvRevision = value
    End Set
  End Property

  Public Property CriteriaSet() As Integer
    Get
      Return mvCriteriaSet
    End Get
    Set(ByVal value As Integer)
      mvCriteriaSet = value
    End Set
  End Property

  Public Property NewCriteriaSet() As Integer
    Get
      Return mvNewCriteriaSet
    End Get
    Set(ByVal value As Integer)
      mvNewCriteriaSet = value
    End Set
  End Property

  Public ReadOnly Property SelectionTable() As String
    Get
      Return mvSelectionTable
    End Get
  End Property

  Public Property OrganisationCriteriaCount() As Integer
    Get
      Return mvOrganisationCriteriaCount
    End Get
    Set(ByVal value As Integer)
      mvOrganisationCriteriaCount = value
    End Set
  End Property

  Public Property DisplayOrgSelection() As Boolean
    Get
      Return mvDisplayOrgSelection
    End Get
    Set(ByVal value As Boolean)
      mvDisplayOrgSelection = value
    End Set
  End Property

  Public ReadOnly Property VariableParameters() As String
    Get
      Return mvVariableParameters
    End Get
  End Property

  Public Property OrganisationMailTo() As String
    Get
      Return mvOrganisationMailTo   'A = AllEmployees, D = DefaultContact, O = Organisation
    End Get
    Set(ByVal value As String)
      mvOrganisationMailTo = value
    End Set
  End Property

  Public Property OrganisationLabelName() As String
    Get
      Return mvOrganisationLabelName
    End Get
    Set(ByVal value As String)
      mvOrganisationLabelName = value
    End Set
  End Property

  Public Property OrganisationMailWhere() As String
    Get
      Return mvOrganisationMailWhere       'O = OrganisationAddress, D = DefaultAddress, U = AddressByUsage
    End Get
    Set(ByVal value As String)
      mvOrganisationMailWhere = value
    End Set
  End Property

  Public Property OrganisationAddressUsage() As String
    Get
      Return mvOrganisationAddressUsage
    End Get
    Set(ByVal value As String)
      mvOrganisationAddressUsage = value
    End Set
  End Property

  Public Property OrganisationRoles() As String
    Get
      Return mvOrganisationRoles
    End Get
    Set(ByVal value As String)
      mvOrganisationRoles = value
    End Set
  End Property

  Public Property IncludeHistoricRoles() As Boolean
    Get
      Return mvIncludeHistoricRoles
    End Get
    Set(ByVal value As Boolean)
      mvIncludeHistoricRoles = value
    End Set
  End Property

  Public Property TaskType() As CareServices.TaskJobTypes
    Get
      Return mvTaskType
    End Get
    Set(ByVal value As CareServices.TaskJobTypes)
      mvTaskType = value
    End Set
  End Property

  Public Property MailingSelection() As Boolean
    Get
      Return mvMailingSelection
    End Get
    Set(ByVal value As Boolean)
      mvMailingSelection = value
    End Set
  End Property

  Public Property CardProductionType() As MembershipCardProductionTypes
    Get
      Return mvCardProductionType
    End Get
    Set(ByVal value As MembershipCardProductionTypes)
      mvCardProductionType = value
    End Set
  End Property

  Public ReadOnly Property SummaryPrintValid() As Boolean
    Get
      Return mvSummaryPrintValid
    End Get
  End Property


  Public Sub Init(ByVal pMailingTypeCode As String, ByVal pCriteriaSet As Integer, Optional ByVal pFromAppealOrSegment As Boolean = False)
    Dim vList As New ParameterList(True)
    vList("ApplicationName") = pMailingTypeCode
    If pCriteriaSet = 0 Then vList("IsSelectionSet") = "Y"
    vList("AppealOrSegment") = CStr(IIf(pFromAppealOrSegment, "Y", "N"))
    mvMailingTypeCode = pMailingTypeCode
    mvCriteriaSet = pCriteriaSet
    Dim vMailingSelection As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtMailingSelections, vList)

    If vMailingSelection IsNot Nothing Then
      If vMailingSelection.Rows.Count > 0 Then
        mvAppealMailing = pFromAppealOrSegment 'BooleanValue(vMailingSelection.Rows(0)("AppealMailing").ToString)
        mvSelectionTable = vMailingSelection.Rows(0)("SelectionTable").ToString
        mvCaption = vMailingSelection.Rows(0)("Caption").ToString
        mvSelectionSet = CInt(vMailingSelection.Rows(0)("SelectionSetNumber"))
        mvMailingTypeCode = vMailingSelection.Rows(0)("MailingTypeCode").ToString
        mvSummaryPrintValid = Convert.ToBoolean(vMailingSelection.Rows(0)("SummaryPrintValid").ToString)
        mvMasterAttribute = vMailingSelection.Rows(0)("MasterAttribute").ToString
        mvDisplayOrgSelection = Convert.ToBoolean(vMailingSelection.Rows(0)("DisplayOrgSelection").ToString)
      End If
    End If

    If vList.Contains("IsSelectionSet") Then vList.Remove("IsSelectionSet")
    If vList.Contains("AppealOrSegment") Then vList.Remove("AppealOrSegment")
    mvMailingType = DataHelper.GetMailingType(vList)
  End Sub

  Public Sub DeleteSelectionSetMailing(ByVal pSelectionSet As Integer, Optional ByVal pDeleteTable As Boolean = False)
    DataHelper.DeleteMailingSelectionSet(pSelectionSet, mvMailingTypeCode, mvRevision, pDeleteTable)
  End Sub

  Public Sub DeleteSelection(ByVal pSetNumber As Integer, ByVal pRevision As Integer, Optional ByVal pDeleteTable As Boolean = False)
    DataHelper.DeleteMailingSelectionSet(pSetNumber, mvMailingTypeCode, pRevision, pDeleteTable)
  End Sub

  Public Function CriteriaContainsORs(ByVal pCriteriaSet As Integer) As Boolean
    Dim vList As New ParameterList(True)
    Dim vCheck As Boolean
    vList("ApplicationName") = mvMailingTypeCode
    vList.IntegerValue("CriteriaSet") = pCriteriaSet
    Dim vParamList As ParameterList = DataHelper.CheckCriteriaContainsORs(vList)
    If vParamList IsNot Nothing Then
      vCheck = Convert.ToBoolean(vParamList("Result"))
    End If
    Return vCheck
  End Function

  Public Function ProcessSelection(ByVal pList As ParameterList) As Integer
    Return DataHelper.ProcessMailingSelection(pList)
  End Function

  Public Function GetMailingSelectionRoughCount(ByVal pCriteriaSet As Integer, Optional ByVal pList As ParameterList = Nothing) As Integer
    If pList Is Nothing Then pList = New ParameterList(True)
    Dim vRoughCount As Integer = 0
    pList("ApplicationName") = mvMailingTypeCode
    pList.IntegerValue("CriteriaSet") = pCriteriaSet
    Dim vParamList As ParameterList = DataHelper.GetMailingSelectionRoughCount(pList)
    If vParamList IsNot Nothing Then
      vRoughCount = vParamList.IntegerValue("Count")
    End If
    Return vRoughCount
  End Function
  Public Function GetMailingSelectionCount(ByVal pSelectionSet As Integer, ByVal pRevision As Integer, ByVal pTypeCode As String) As Integer
    Dim vList As New ParameterList(True)
    Dim vRoughCount As Integer = 0
    vList.IntegerValue("Revision") = pRevision
    vList.IntegerValue("SelectionSetNumber") = pSelectionSet
    vList("ApplicationCode") = pTypeCode
    Dim vParamList As ParameterList = DataHelper.GetMailingSelectionCount(vList)
    If vParamList IsNot Nothing Then
      vRoughCount = vParamList.IntegerValue("Count")
    End If
    Return vRoughCount
  End Function

  Public Sub ProcessMailingCriteriaWithOptional(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByVal pProcessVariables As Boolean, ByVal pEditSegmentCriteria As Boolean, ByRef pList As ParameterList, ByRef pSuccess As Boolean)
    ProcessMailingCriteria(pCriteriaSet, pProcessVariables, pEditSegmentCriteria, pList, pSuccess)
    If pList IsNot Nothing AndAlso pList.Count > 0 AndAlso pMailingSelection IsNot Nothing Then
      If pList.Contains("OrgMailTo") Then pMailingSelection.OrganisationMailTo = pList("OrgMailTo")
      If pList.Contains("OrgMailWhere") Then
        If pList.Item("OrgMailWhere") <> "U" Then pMailingSelection.OrganisationAddressUsage = ""
        pMailingSelection.OrganisationMailWhere = pList("OrgMailWhere")
      End If
      If pList.Contains("OrgAddressUsage") Then pMailingSelection.OrganisationAddressUsage = pList("OrgAddressUsage")
      If pList.Contains("OrgIncludeHistoricRoles") Then pMailingSelection.IncludeHistoricRoles = CBool(IIf(pList("OrgIncludeHistoricRoles") = "Y", True, False))
      If pList.Contains("OrgLabelName") Then pMailingSelection.OrganisationLabelName = pList("OrgLabelName")
      If pList.Contains("OrgRoles") Then pMailingSelection.OrganisationRoles = pList("OrgRoles")
    End If
  End Sub

  Public Sub ProcessMailingCriteria(ByVal pCriteriaSet As Integer, ByVal pProcessVariables As Boolean, ByVal pEditSegmentCriteria As Boolean, ByRef pList As ParameterList, ByRef pSuccess As Boolean)
    pSuccess = True
    If pList Is Nothing Then pList = New ParameterList(True)
    Dim vPanelItems As PanelItems = Nothing
    If pCriteriaSet > 0 Then pList.IntegerValue("CriteriaSetNumber") = pCriteriaSet 'mvMailingInfo.CriteriaSet
    pList("ApplicationName") = Me.MailingTypeCode         'Needs the code from the mailing type for campaign mailings  AppValues.MailingApplicationCode(Me.TaskType) 'mvMailingInfo.MailingTypeCode
    Dim vDataSet As DataSet = DataHelper.GetCampaignCriteriaVariableControls(pList, Nothing)
    Dim vControls As New PanelItems("CriteriaVariables")    'Add variables controls to PanelItems if any
    If vDataSet IsNot Nothing Then
      Dim vRow As DataRow = Nothing
      Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
      If vTable IsNot Nothing Then
        For Each vRow In vTable.Rows
          Dim vPanelItem As New PanelItem(vRow)
          vControls.Add(vPanelItem)
        Next
      End If
      If vDataSet.Tables.Contains("Parameters1") OrElse vDataSet.Tables.Contains("Parameters") Then          'Add additional parameters for Organisation Selection if any
        If vDataSet.Tables.Contains("Parameters1") Then
          vRow = vDataSet.Tables("Parameters1").Rows(0)
        Else
          vRow = vDataSet.Tables("Parameters").Rows(0)
        End If
        pList.Add("OrganisationCriteriaCount", vRow("OrganisationCriteriaCount").ToString)
        pList.Add("ContactCriteriaCount", vRow("ContactCriteriaCount").ToString)
        If vRow.Table.Columns.Contains("OrgMailTo") Then pList.Add("OrgMailTo", vRow("OrgMailTo").ToString)
        If vRow.Table.Columns.Contains("OrgMailWhere") Then pList.Add("OrgMailWhere", vRow("OrgMailWhere").ToString)
        If vRow.Table.Columns.Contains("OrgRoles") Then pList.Add("OrgRoles", vRow("OrgRoles").ToString)
        If vRow.Table.Columns.Contains("OrgAddressUsage") Then pList.Add("OrgAddressUsage", vRow("OrgAddressUsage").ToString)
        If vRow.Table.Columns.Contains("OrgLabelName") Then pList.Add("OrgLabelName", vRow("OrgLabelName").ToString)
      End If
      vPanelItems = vControls
    End If

    If vPanelItems IsNot Nothing AndAlso vPanelItems.Count > 0 Then
      If pProcessVariables Then
        Dim vAppParameters As frmApplicationParameters
        Dim vVariableList As ParameterList
        vAppParameters = New frmApplicationParameters(EditPanelInfo.OtherPanelTypes.optCriteriaVariables, vPanelItems, Nothing, "")
        If vAppParameters.ShowDialog(CurrentMainForm) = System.Windows.Forms.DialogResult.OK Then

          vVariableList = vAppParameters.ReturnList
          If vVariableList IsNot Nothing AndAlso vVariableList.Count > 0 Then
            For Each vValue As DictionaryEntry In vVariableList
              If Not pList.Contains(vValue.Key.ToString) Then pList.Add(vValue.Key.ToString, vValue.Value.ToString)
            Next
          Else
            pSuccess = False
          End If
        Else
          pSuccess = False
        End If
      End If
    End If
    If pSuccess AndAlso pList.Contains("OrganisationCriteriaCount") AndAlso pList.IntegerValue("OrganisationCriteriaCount") > 0 AndAlso Me.DisplayOrgSelection Then
      Dim vOrgSelectList As ParameterList
      Dim vPassList As ParameterList = Nothing
      If pList.IntegerValue("ContactCriteriaCount") > 0 Then
        vOrgSelectList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optAddressSelectionOptions, Nothing, vPassList, "")
      Else
        vPassList = New ParameterList()
        vPassList("MailingApplicationCode") = Me.MailingTypeCode
        If Me.AppealMailing Then
          vPassList("ApplicationName") = "CA"
        Else
          vPassList("ApplicationName") = Me.MailingTypeCode
        End If
        vOrgSelectList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optOrganisationSelectionOptions, Nothing, vPassList, "")
      End If
      pSuccess = FormHelper.AddOrgMailingParameters(vOrgSelectList, pList)
    End If
  End Sub

End Class
