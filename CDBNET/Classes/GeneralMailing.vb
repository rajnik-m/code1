Public Class GeneralMailing
  Private mvMailingInfo As MailingInfo
  Private WithEvents mvfrmGenMGen As frmGenMGen
  Private WithEvents mvfrmGenMMerge As frmGenMMerge
  Private WithEvents mvFrmEditCriteria As frmEditCriteria
  Private mvSelectionSet As Integer
  Private mvOptActions As Boolean
  Private mvOptActivityQty As Boolean
  Private mvOptCustomData As Boolean
  Private mvOptContactGroups As Boolean
  Private mvOptOrgGroups As Boolean
  Private mvOptIntranet As Boolean
  Private mvOptMeetings As Boolean
  Private mvOptEvents As Boolean
  Private mvOptManualCache As Boolean
  Private mvCriteriaSet As Integer
  Private mvMailingTypeCode As String = ""

  Public Enum SaveTypes
    CriteriaSet
    List
  End Enum

  Public Sub New(ByVal pMailingTypeCode As String)
    mvMailingInfo = New MailingInfo()
    initialise()
  End Sub
  Public Sub New(ByVal pMailingType As CareNetServices.MailingTypes, ByVal pTaskType As CareServices.TaskJobTypes)
    mvMailingInfo = New MailingInfo()
    mvMailingTypeCode = AppValues.MailingApplicationCode(pTaskType)
    mvMailingInfo.TaskType = pTaskType
    initialise()
  End Sub

  Public ReadOnly Property MailingInfo As MailingInfo
    Get
      Return mvMailingInfo
    End Get
  End Property

  Private Sub initialise()
    Dim vMailingTypeCode As String
    vMailingTypeCode = "GM"

    If mvMailingTypeCode <> "" Then vMailingTypeCode = mvMailingTypeCode

    mvMailingInfo.Init(vMailingTypeCode, mvSelectionSet)
    mvSelectionSet = mvMailingInfo.SelectionSet
    mvOptCustomData = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_custom_data, False)
    mvOptContactGroups = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_contact_groups, False)
    mvOptOrgGroups = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_organisation_groups, False)
    mvOptManualCache = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.option_cache_manual_update, False)
  End Sub

  Public Sub Process(ByVal pCriteriaSet As Integer)
    Dim vDisplaySchedule As Boolean
    Dim vResultDialog As DialogResult
    Dim vList As New ParameterList(True)
    Select Case mvMailingInfo.MailingType
      Case CareNetServices.MailingTypes.mtyEventBookings, CareNetServices.MailingTypes.mtyEventPersonnel, _
           CareNetServices.MailingTypes.mtyEventAttendees, CareNetServices.MailingTypes.mtyEventBookings, _
           CareNetServices.MailingTypes.mtyEventSponsors, CareNetServices.MailingTypes.mtyIrishGiftAid, _
           CareNetServices.MailingTypes.mtyExamBookings, CareNetServices.MailingTypes.mtyExamCandidates
        vDisplaySchedule = False
      Case Else
        vDisplaySchedule = True
    End Select
    If Not vDisplaySchedule Then
      vResultDialog = DialogResult.No
    Else
      vResultDialog = FormHelper.ScheduleTask(vList)

    End If
    Select Case vResultDialog
      Case DialogResult.Yes
        'run a task when Schedule
        FormHelper.ProcessTask(mvMailingInfo.TaskType, vList, True, FormHelper.ProcessTaskScheduleType.ptsAlwaysSchedule)
      Case DialogResult.No
        If pCriteriaSet > 0 Then
          mvCriteriaSet = pCriteriaSet
        End If
        mvMailingInfo.SelectionSet = mvSelectionSet
        mvMailingInfo.CriteriaSet = mvCriteriaSet
        mvMailingInfo.Revision = 0
        mvFrmEditCriteria = New frmEditCriteria(mvMailingInfo, "Test")
        mvFrmEditCriteria.ShowDialog()
        Dim vParams As New ParameterList(True)
        If pCriteriaSet = 0 Then
          If mvMailingInfo.CriteriaSet > 0 Then
            vParams("CriteriaSet") = mvMailingInfo.CriteriaSet.ToString
            DataHelper.DeleteCriteriaSetDetails(vParams)
          End If
        Else
          If mvMailingInfo.ExclusionCriteriaSet > 0 Then
            vParams("CriteriaSet") = mvMailingInfo.ExclusionCriteriaSet.ToString
            DataHelper.DeleteCriteriaSetDetails(vParams)
          End If
        End If
        mvMailingInfo.DeleteSelection(mvSelectionSet, 0, True)
    End Select

  End Sub

  Public Sub ShowGeneralMailingForm()
    mvfrmGenMGen = New frmGenMGen(mvMailingTypeCode, mvMailingInfo, mvSelectionSet)
    mvfrmGenMGen.ShowDialog()
  End Sub

  Private Function GetMailingSelectionCount() As Integer
    'Return the number of selected records
    Dim vSelectedRecords As Integer
    vSelectedRecords = mvMailingInfo.GetMailingSelectionCount(mvMailingInfo.SelectionSet, mvMailingInfo.Revision, mvMailingTypeCode)
    ShowInformationMessage(InformationMessages.ImCOntactSelected, vSelectedRecords.ToString)     '%s contacts selected
    Return vSelectedRecords
  End Function
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pRunPhase"></param>
  ''' <param name="pList"></param>
  ''' <remarks>BR19394 - If you change this event handler, check corresponding event handler in frmTableMaintenance, it is a copy of this event handler </remarks>
  Private Sub mvFrmEditCriteria_ProcessSelection(ByVal pRunPhase As String, ByVal pList As ParameterList) Handles mvFrmEditCriteria.ProcessSelection
    Try
      Dim vStoreGvRev As Integer
      If pList Is Nothing Then pList = New ParameterList(True)
      Dim vCountOnly As Boolean

      mvMailingInfo.GenerateStatus = MailingInfo.MailingGenerateResult.mgrNone
      mvCriteriaSet = mvMailingInfo.CriteriaSet
      mvSelectionSet = mvMailingInfo.SelectionSet
      vCountOnly = (pRunPhase = "Count")
      If mvMailingInfo.CriteriaRows = 0 Then
        If mvMailingInfo.Revision > 0 Then
          If mvMailingInfo.SelectionSet > 0 Then mvMailingInfo.SelectionCount = GetMailingSelectionCount()
          If mvMailingInfo.SelectionCount > 0 Then
            mvfrmGenMGen = New frmGenMGen(mvMailingTypeCode, mvMailingInfo, mvSelectionSet)
            mvfrmGenMGen.ShowDialog()
          End If
        Else
          ShowInformationMessage(InformationMessages.ImNoCriteria)    'No criteria entered
        End If
      Else
        If mvMailingInfo.Revision = 0 Then
          'There is no selection set
          vStoreGvRev = mvMailingInfo.Revision
          mvMailingInfo.Revision = mvMailingInfo.Revision + 1
          pList("SelectionSetNumber") = mvSelectionSet.ToString
          pList("Revision") = mvMailingInfo.Revision.ToString
          pList("ApplicationName") = AppValues.MailingApplicationCode(mvMailingInfo.TaskType) 'mvMailingInfo.MailingTypeCode.ToString
          pList("RunPhase") = pRunPhase
          pList("CriteriaSet") = mvCriteriaSet.ToString
          pList("ExclusionCriteria") = mvMailingInfo.ExclusionCriteriaSet.ToString
          pList("OrgMailTo") = mvMailingInfo.OrganisationMailTo
          pList("OrgMailWhere") = mvMailingInfo.OrganisationMailWhere
          pList("OrgLabelName") = mvMailingInfo.OrganisationLabelName
          pList("OrgAddressUsage") = mvMailingInfo.OrganisationAddressUsage
          pList("OrgRoles") = mvMailingInfo.OrganisationRoles
          pList("OrgIncludeHistoricRoles") = IIf(mvMailingInfo.IncludeHistoricRoles, "Y", "N").ToString
          pList("BypassCount") = IIf(mvMailingInfo.BypassCriteriaCount, "Y", "N").ToString
          pList("GeneralMailing") = "Y"

          If vCountOnly Then
            Dim vResults As ParameterList = DataHelper.ProcessMailingCount(pList)
            If vResults.Contains("MailingCount") Then mvMailingInfo.SelectionCount = vResults.IntegerValue("MailingCount") 'GetMailingSelectionCount()
            mvMailingInfo.Revision = 0
          Else
            DataHelper.GetMailingFile(pList, "", False, True)
          End If
          If mvMailingInfo.Revision = 0 Then
            'MouseNormal()
            mvMailingInfo.Revision = vStoreGvRev
          Else
            mvMailingInfo.SelectionCount = GetMailingSelectionCount()
            'MouseNormal()
            If mvMailingInfo.SelectionCount > 0 Then
              mvfrmGenMGen = New frmGenMGen(mvMailingTypeCode, mvMailingInfo, mvSelectionSet)
              mvfrmGenMGen.ShowDialog()
            Else
              mvMailingInfo.Revision = vStoreGvRev
            End If
          End If
        Else
          Dim vForm As frmGenMMerge = New frmGenMMerge(mvMailingInfo, pList)
          vForm.ShowDialog()
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub mvFrmEditCriteria_SaveCriteria() Handles mvFrmEditCriteria.SaveCriteria
    Try
      'Get a new control number for a criteria set and
      'add a new entry to the criteria sets table
      ProcessSave(SaveTypes.CriteriaSet)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ProcessSave(ByVal pSaveType As SaveTypes)
    Dim vParams As New ParameterList(True)
    Dim vCaption As String = "Save"
    Dim vSaveResults As ParameterList
    vParams("ApplicationName") = mvMailingInfo.MailingTypeCode
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
    If vSaveResults IsNot Nothing Then ProcessSave(vSaveResults)
  End Sub

  Private Sub ProcessSave(ByVal pList As ParameterList)

    If pList("CriteriaSetDesc").ToString.Length > 0 Then
      Dim vParams As New ParameterList(True)
      Dim vCriteriaSetNumber As Integer = 0

      Dim vDataTable As DataTable = Nothing
      If mvMailingInfo.CriteriaSet > 0 Then
        vParams.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
        vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCriteriaSetDetails, vParams)
      End If

      Dim vCriteriaSet As New ParameterList(True)
      vCriteriaSet.IntegerValue("CriteriaSetNumber") = vCriteriaSetNumber
      vCriteriaSet("UserName") = pList("UserName").ToString
      vCriteriaSet("Department") = DataHelper.UserInfo.Department
      vCriteriaSet("CriteriaSetDesc") = pList("CriteriaSetDesc").ToString
      vCriteriaSet("ApplicationName") = mvMailingInfo.MailingTypeCode
      If pList.ContainsKey("ReportCode") AndAlso pList("ReportCode").ToString.Length > 0 Then vCriteriaSet("ReportCode") = pList("ReportCode").ToString
      If pList.ContainsKey("StandardDocument") AndAlso pList("StandardDocument").ToString.Length > 0 Then vCriteriaSet("StandardDocument") = pList("StandardDocument").ToString
      Dim vReturnList As ParameterList = DataHelper.AddCriteriaSet(vCriteriaSet)
      vCriteriaSetNumber = vReturnList.IntegerValue("CriteriaSetNumber")

      If vDataTable IsNot Nothing Then
        For Each vRow As DataRow In vDataTable.Rows
          Dim vCriteriaSetValues As New ParameterList(True)
          vCriteriaSetValues.IntegerValue("CriteriaSet") = vCriteriaSetNumber
          vCriteriaSetValues("SequenceNumber") = vRow.Item("SequenceNumber").ToString
          vCriteriaSetValues("SearchArea") = vRow.Item("SearchArea").ToString
          vCriteriaSetValues("IE") = vRow.Item("IE").ToString
          vCriteriaSetValues("CO") = vRow.Item("CO").ToString
          vCriteriaSetValues("MainValue") = vRow.Item("MainValue").ToString
          vCriteriaSetValues("SubsidiaryValue") = vRow.Item("SubsidiaryValue").ToString
          vCriteriaSetValues("Period") = vRow.Item("Period").ToString
          If vRow.Item("Counted").ToString <> "" AndAlso CInt(vRow.Item("Counted").ToString) > 0 Then vCriteriaSetValues("Counted") = vRow.Item("Counted").ToString
          vCriteriaSetValues("AndOr") = vRow.Item("AndOr").ToString
          vCriteriaSetValues("LeftParenthesis") = vRow.Item("LeftParenthesis").ToString
          vCriteriaSetValues("RightParenthesis") = vRow.Item("RightParenthesis").ToString
          Dim vResult As ParameterList = DataHelper.AddCriteriaSetDetails(vCriteriaSetValues)
        Next
      End If
    End If
  End Sub

  Public Sub ProcessMailingCriteria(ByVal pCriteriaSet As Integer, ByRef pList As ParameterList, ByRef pSuccess As Boolean)
    mvMailingInfo.ProcessMailingCriteria(pCriteriaSet, True, False, pList, pSuccess)
  End Sub
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pMailingSelection"></param>
  ''' <param name="pCriteriaSet"></param>
  ''' <param name="pSuccess"></param>
  ''' <remarks>BR19394 - If you change this event handler, check corresponding event handler in frmTableMaintenance, it is a copy of this event handler</remarks>
  Private Sub mvFrmEditCriteria_ProcessMailingCriteria(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByRef pSuccess As Boolean) Handles mvFrmEditCriteria.ProcessMailingCriteria
    mvMailingInfo.ProcessMailingCriteria(pCriteriaSet, True, False, Nothing, pSuccess)
  End Sub
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="pMailingSelection"></param>
  ''' <param name="pCriteriaSet"></param>
  ''' <param name="pProcessVariables"></param>
  ''' <param name="pEditSegmentCriteria"></param>
  ''' <param name="pList"></param>
  ''' <param name="pSuccess"></param>
  ''' <remarks>BR19394 - If you change this event handler, check corresponding event handler in frmTableMaintenance, it is a copy of this event handler</remarks>
  Private Sub mvFrmEditCriteria_ProcessMailingCriteriaWithOptional(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByVal pProcessVariables As Boolean, ByVal pEditSegmentCriteria As Boolean, ByRef pList As ParameterList, ByRef pSuccess As Boolean) Handles mvFrmEditCriteria.ProcessMailingCriteriaWithOptional
    mvMailingInfo.processMailingCriteriaWithOptional(pMailingSelection, pCriteriaSet, pProcessVariables, pEditSegmentCriteria, pList, pSuccess)
  End Sub
End Class
