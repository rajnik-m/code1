Imports CDBNETCL.ExamsAccess
Imports CDBNETCL.CareNetServices
Imports CDBNETCL.ExamSelector
Imports System.Globalization
Imports CDBNETBiz

Public Class frmExams
  Implements IPanelVisibility

#Region "IPanelVisibility"

  Public Sub SetPanelVisibility() Implements IPanelVisibility.SetPanelVisibility
    'If mvHeader IsNot Nothing AndAlso mvEventInfo IsNot Nothing AndAlso mvEventInfo.EventNumber > 0 Then RefreshHeader()
    splBottom.Panel1Collapsed = Not MainHelper.ShowSelectionPanel
  End Sub

  Public Property PanelHasFocus() As Boolean Implements IPanelVisibility.PanelHasFocus
    Get
      Return sel.Focused
    End Get
    Set(ByVal value As Boolean)
      If value Then
        sel.Focus()
      Else
        Me.Focus()
      End If
    End Set
  End Property

#End Region

#Region "Maintenance Parent Form Methods"

  Public Overrides ReadOnly Property SizeMaintenanceForm() As Boolean
    Get
      Return True
    End Get
  End Property

  Public Overrides Sub RefreshData(ByVal pType As CareServices.XMLMaintenanceControlTypes)
    'Do Nothing
    RefreshCard()
  End Sub

  Public Overrides Sub RefreshData()
    'BR 21694 - User gets the Object Reference error when try's to refresh the Examination data
    If cboSessions.SelectedValue IsNot Nothing Then
      sel.Init(mvSelectorType, IntegerValue(cboSessions.SelectedValue))
      sel.Focus()
    End If
  End Sub

#End Region

#Region "Enums and class variables"

  Private Const IMAGEBUTTONGAP As Integer = 6
  Private Const IMAGEBUTTONHEIGHT As Integer = 32
  Private Const IMAGEBUTTONTOP As Integer = 4

  Private mvSelectorType As ExamSelector.SelectionType = ExamSelector.SelectionType.Courses
  Private mvExamDataType As ExamsAccess.XMLExamDataSelectionTypes
  Private mvExamMaintenanceType As ExamsAccess.XMLExamMaintenanceTypes
  Private WithEvents mvExamScheduleSelector As ExamSelector
  Private mvExamExemptionId As Integer
  Private mvExamUnitId As Integer
  Private mvExamUnitLinkId As Integer
  Private mvExamPersonnelId As Integer
  Private mvExamCentreId As Integer
  Private mvExamSessionId As Integer
  Private mvSelectedRow As Integer
  Private mvRefreshSelector As Boolean
  Private mvParentID As Integer
  Private mvParentLinkID As Integer
  Private mvCentreOrganisation As Integer
  Private mvCourseSessionId As Integer
  Private mvSelectColumnName As String = "Select"
  Private mvActionNumber As Integer = 0
  Private mvExamCentreUnitId As Integer
  Private mvExamSelectorItem As ExamSelectorItem
  Private WithEvents mvExamsMenu As New ExamsMenu(Me)
  Private WithEvents mvSelExamsMenu As New ExamsMenu(Me)
  Private WithEvents mvCustomiseMenu As CustomiseMenu
  Private WithEvents mvButtonsCustomiseMenu As ExamsMenu
  Private WithEvents mvExamsGridMenu As New ExamsMenu(Me)
  Private WithEvents mvDocumentMenu As BaseDocumentMenu = Nothing
  Private WithEvents mvActionMenu As New ActionMenu(Me)
  Private mvDocumentsDataSet As DataSet
  Private mvDocLinkDataSource As DataSet
  Private mvDocAnalysisDataSource As DataSet
  Private mvPanelProportions As New Dictionary(Of XMLExamDataSelectionTypes, Double)
  Private mvStudyModeList As List(Of String) = Nothing
  Private mvCentreContact As Integer

#End Region

  Private WithEvents mvCurrentEpl As EditPanel = Nothing

#Region "Contructors and Initialisation"

  Public Sub New()
    ' This call is required by the designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
    mvActionMenu.ActionType = ActionMenu.ActionTypes.ExamActions
  End Sub

  Private Sub InitialiseControls()
    Me.splTop.Panel2.SuspendLayout()
    Me.splTop.SuspendLayout()
    Me.splBottom.Panel1.SuspendLayout()
    Me.splBottom.Panel2.SuspendLayout()
    Me.splBottom.SuspendLayout()
    Me.pnlCommands.SuspendLayout()
    Me.tabMaster.SuspendLayout()
    Me.tabMain.SuspendLayout()
    Me.splRight.Panel1.SuspendLayout()
    Me.splRight.Panel2.SuspendLayout()
    Me.splRight.SuspendLayout()
    Me.bpl.SuspendLayout()
    Me.tabCustomForm.SuspendLayout()
    Me.dgrMenuStrip.SuspendLayout()
    Me.dgr2MenuStrip.SuspendLayout()
    Me.dplMenuStrip.SuspendLayout()
    Me.dgr1MenuStrip.SuspendLayout()
    Me.dgr0MenuStrip.SuspendLayout()
    Me.SuspendLayout()
    mvCurrentEpl = epl
    mvExamSelectorItem = New ExamSelectorItem(Nothing, SelectionType.Courses)
    SetControlTheme()
    SettingsName = "Exams_CardSet"
    Me.Text = "Examinations Maintenance"
    Me.splTop.Panel1Collapsed = True
    dgr.MaxGridRows = DisplayTheme.DefaultMaxGridRows
    dgr.AutoSetHeight = True
    mvCustomiseMenu = New CustomiseMenu
    mvButtonsCustomiseMenu = New ExamsMenu(Me)
    mvButtonsCustomiseMenu.Customise = True
    PopulateSessions()
    sel.TreeContextMenu = mvExamsMenu
    selExams.TreeContextMenu = mvSelExamsMenu
    sel.ExamMaintenance = True
    sel.NodesCustomiseMenu = New ExamsMenu(Me)
    SetupImageButtons()
    sel.Init(mvSelectorType)
    tabMaster.Appearance = TabAppearance.FlatButtons
    tabMaster.ItemSize = New Size(0, 1)
    tabMaster.SizeMode = TabSizeMode.Fixed
    tabMaster.Padding = New Point(0, 0)
    tabMaster.Margin = New Padding(0)
    Me.splTop.Panel2.ResumeLayout()
    Me.splTop.ResumeLayout()
    Me.splBottom.Panel1.ResumeLayout()
    Me.splBottom.Panel2.ResumeLayout()
    Me.splBottom.ResumeLayout()
    Me.pnlCommands.ResumeLayout()
    Me.tabMaster.ResumeLayout()
    Me.tabMain.ResumeLayout()
    Me.splRight.Panel1.ResumeLayout()
    Me.splRight.Panel2.ResumeLayout()
    Me.splRight.ResumeLayout()
    Me.bpl.ResumeLayout()
    Me.tabCustomForm.ResumeLayout()
    Me.dgrMenuStrip.ResumeLayout()
    Me.dgr2MenuStrip.ResumeLayout()
    Me.dplMenuStrip.ResumeLayout()
    Me.dgr1MenuStrip.ResumeLayout()
    Me.dgr0MenuStrip.ResumeLayout()
    Me.ResumeLayout()
  End Sub

  Private Sub SetupImageButtons()
    'Get AccessControl to set visibility
    Dim vACDT As DataTable = Nothing
    Dim vDT As DataTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtUserMenuAccess).Copy   'Don't want to set the filter for everything
    If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
      Dim vFilter As String = "Item Like 'SCXMB%'"
      vDT.DefaultView.RowFilter = vFilter
      vACDT = vDT.DefaultView.ToTable()
    End If

    'Manipulate the Image Buttons using both the DisplayListMaintenance and AccessControl DataTable's
    Dim vBtnCount As Integer = 0
    Dim vFirstButton As ImageButton = Nothing
    Dim vList As New ParameterList(True, True)
    vList("ExamMaintenance") = "Y"
    vList.IntegerValue("ExamDataType") = CInt(XMLExamDataSelectionTypes.ExamMaintenanceButtons)
    vList("ContactGroup") = "CON"
    vDT = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSelectionPages, vList, False)    'Cannot cache this otherwise we cannot get customised data back
    If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
      'Make all buttons invisible first, then reset as required
      imgCourse.Visible = False
      imgCentres.Visible = False
      imgExemptions.Visible = False
      imgPersonnel.Visible = False
      imgSessions.Visible = False
      Dim vButtonTop As Integer = IMAGEBUTTONTOP
      Dim vDesc As String = String.Empty
      Dim vReadOnly As Boolean = False
      Dim vVisible As Boolean = True
      For Each vDLRow As DataRow In vDT.Rows
        vDesc = vDLRow.Item("Description").ToString
        vReadOnly = BooleanValue(vDLRow.Item("ReadOnly").ToString)
        vVisible = True
        Select Case vDLRow.Item("Code").ToString
          Case "Courses"
            vVisible = GetImageButtonVisibility(vACDT, imgCourse.Tag.ToString.Substring(0, 6))
            SetImageButton(imgCourse, vDesc, vButtonTop, vVisible, vReadOnly)
            If vVisible AndAlso vFirstButton Is Nothing Then vFirstButton = imgCourse
            If vVisible Then vBtnCount += 1
          Case "Centres"
            vVisible = GetImageButtonVisibility(vACDT, imgCentres.Tag.ToString.Substring(0, 6))
            SetImageButton(imgCentres, vDesc, vButtonTop, vVisible, vReadOnly)
            If vVisible AndAlso vFirstButton Is Nothing Then vFirstButton = imgCentres
            If vVisible Then vBtnCount += 1
          Case "Exemptions"
            vVisible = GetImageButtonVisibility(vACDT, imgExemptions.Tag.ToString.Substring(0, 6))
            SetImageButton(imgExemptions, vDesc, vButtonTop, vVisible, vReadOnly)
            If vVisible AndAlso vFirstButton Is Nothing Then vFirstButton = imgExemptions
            If vVisible Then vBtnCount += 1
          Case "Personnel"
            vVisible = GetImageButtonVisibility(vACDT, imgPersonnel.Tag.ToString.Substring(0, 6))
            SetImageButton(imgPersonnel, vDesc, vButtonTop, vVisible, vReadOnly)
            If vVisible AndAlso vFirstButton Is Nothing Then vFirstButton = imgPersonnel
            If vVisible Then vBtnCount += 1
          Case "Sessions"
            vVisible = GetImageButtonVisibility(vACDT, imgSessions.Tag.ToString.Substring(0, 6))
            SetImageButton(imgSessions, vDesc, vButtonTop, vVisible, vReadOnly)
            If vVisible AndAlso vFirstButton Is Nothing Then vFirstButton = imgSessions
            If vVisible Then vBtnCount += 1
        End Select
        If vVisible Then vButtonTop += (IMAGEBUTTONHEIGHT + IMAGEBUTTONGAP)
        If vBtnCount = 1 AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDisplayListMaintenance) Then
          'Only add the customise menu to the top button
          mvButtonsCustomiseMenu.ExamDataType = XMLExamDataSelectionTypes.ExamMaintenanceButtons
          mvButtonsCustomiseMenu.DataSelectionType = IntegerValue(vDLRow.Item("DataSelection").ToString)
          RemoveHandler mvButtonsCustomiseMenu.CustomiseMenu, AddressOf CustomiseButtons
          AddHandler mvButtonsCustomiseMenu.CustomiseMenu, AddressOf CustomiseButtons
          If vFirstButton IsNot Nothing Then vFirstButton.ContextMenuStrip = mvButtonsCustomiseMenu
        End If
      Next
    Else
      vFirstButton = imgCourse
    End If

    If vBtnCount > 0 Then
      If vBtnCount < 5 Then
        Dim vPanelHeight As Integer = (IMAGEBUTTONHEIGHT * vBtnCount)
        vPanelHeight += (IMAGEBUTTONGAP * vBtnCount)
        vPanelHeight += 5
        pnlCommands.Height = vPanelHeight
      End If
      If Not (vFirstButton Is imgCourse) Then
        If vFirstButton Is imgPersonnel Then
          mvSelectorType = SelectionType.Personnel
        ElseIf vFirstButton Is imgCentres Then
          mvSelectorType = SelectionType.Centres
        ElseIf vFirstButton Is imgSessions Then
          mvSelectorType = SelectionType.Sessions
        Else
          mvSelectorType = SelectionType.Exemptions
        End If
      End If
    Else
      Throw New CareException(CareException.ErrorNumbers.enExamMaintenanceIncorrectSetup)
    End If
    cboSessions.Visible = (mvSelectorType = SelectionType.Courses)
    imgExemptions.DrawBorder = mvSelectorType = SelectionType.Exemptions
    imgPersonnel.DrawBorder = mvSelectorType = SelectionType.Personnel
    imgCourse.DrawBorder = mvSelectorType = SelectionType.Courses
    imgCentres.DrawBorder = mvSelectorType = SelectionType.Centres
    imgSessions.DrawBorder = mvSelectorType = SelectionType.Sessions

  End Sub

  ''' <summary>Set ImageButton properties based upon AccessControl and DisplayListMaintenance.</summary>
  Private Sub SetImageButton(ByVal pButton As ImageButton, ByVal pText As String, ByVal pTop As Integer, ByVal pVisible As Boolean, ByVal pReadOnly As Boolean)
    With pButton
      .Text = pText
      .Top = pTop
      .Visible = pVisible
      If .Tag.ToString.EndsWith("RO") Then .Tag = .Tag.ToString.Substring(0, (.Tag.ToString.Length - 3))
      If pReadOnly Then .Tag = .Tag.ToString & "_RO"
    End With
  End Sub

  ''' <summary>Get visibility of image Button from user's access control items.</summary>
  ''' <param name="pAccessControlDT">The UserMenuAccess DataTable</param>
  ''' <param name="pCode">Image button code.</param>
  ''' <returns>True if button is visible, otherwise False.</returns>
  Private Function GetImageButtonVisibility(ByVal pAccessControlDT As DataTable, ByVal pCode As String) As Boolean
    Dim vVisible As Boolean = True
    If pAccessControlDT IsNot Nothing AndAlso pAccessControlDT.Rows.Count > 0 Then
      For Each vRow As DataRow In pAccessControlDT.Rows
        If vRow.Item("Item").ToString.ToUpper = pCode.ToUpper Then vVisible = BooleanValue(vRow.Item("Visible").ToString)
      Next
    End If
    Return vVisible
  End Function

  Private Sub RefreshCard()
    Try
      Me.splTop.Panel2.SuspendLayout()
      Me.splTop.SuspendLayout()
      Me.splBottom.Panel1.SuspendLayout()
      Me.splBottom.Panel2.SuspendLayout()
      Me.splBottom.SuspendLayout()
      Me.pnlCommands.SuspendLayout()
      Me.tabMaster.SuspendLayout()
      Me.tabMain.SuspendLayout()
      Me.splRight.Panel1.SuspendLayout()
      Me.splRight.Panel2.SuspendLayout()
      Me.splRight.SuspendLayout()
      Me.bpl.SuspendLayout()
      Me.tabCustomForm.SuspendLayout()
      Me.dgrMenuStrip.SuspendLayout()
      Me.dgr2MenuStrip.SuspendLayout()
      Me.dplMenuStrip.SuspendLayout()
      Me.dgr1MenuStrip.SuspendLayout()
      Me.dgr0MenuStrip.SuspendLayout()
      Me.SuspendLayout()

      Dim vDataSet As DataSet
      Dim vShowGrid As Boolean = True
      epl.Visible = False
      Me.dspTabGrid.Visible = False
      Me.dspTabGrid.SetText("")
      UpperEditPanel.Visible = False
      dpl.Visible = False
      splTab.Visible = False
      Dim vShowPanel As Boolean = False
      mvDocumentsDataSet = Nothing

      If Me.dspTabGrid.DisplayGrid(0).ContextMenuStrip IsNot Nothing Then
        Me.dspTabGrid.DisplayGrid(0).ContextMenuStrip = Nothing
        Me.dspTabGrid.DisplayGrid(0).SetToolBarVisible()
      End If
      If Me.dspTabGrid.DisplayGrid(1).ContextMenuStrip IsNot Nothing Then
        Me.dspTabGrid.DisplayGrid(1).ContextMenuStrip = Nothing
        Me.dspTabGrid.DisplayGrid(1).SetToolBarVisible()
      End If

      If Me.dgr.ContextMenuStrip IsNot Nothing Then
        dgr.ContextMenuStrip = Nothing
        dgr.SetToolBarVisible()
      End If

      If mvDocumentMenu IsNot Nothing Then mvDocumentMenu = Nothing
      dgr0MenuStrip.Visible = False
      dgr1MenuStrip.Visible = False

      If mvExamDataType = XMLExamDataSelectionTypes.ExamCentres Or
         mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitDetails Or
         mvExamDataType = XMLExamDataSelectionTypes.ExamUnits Then 'fixed bug introduced in Changeset 167702
        mvCurrentEpl = Me.UpperEditPanel
      Else
        mvCurrentEpl = Me.epl
      End If
      tabMaster.SelectedTab = tabMain
      Dim vList As New ParameterList(True, True)
      Dim vItemID As Integer
      Select Case mvExamDataType
        Case XMLExamDataSelectionTypes.ExamCentreAssessmentTypes
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreAssessmentTypes
          vList.IntegerValue("ExamCentreId") = mvExamCentreId
        Case XMLExamDataSelectionTypes.ExamCentres
          vList("SystemColumns") = "N"
          vShowGrid = False
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentres
          vList.IntegerValue("ExamCentreId") = mvExamCentreId
          vItemID = mvExamCentreId
        Case XMLExamDataSelectionTypes.ExamCentreContacts
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreContacts
          vList.IntegerValue("ExamCentreId") = mvExamCentreId
        Case XMLExamDataSelectionTypes.ExamCentreActions
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreActions
          vList.IntegerValue("ExamCentreId") = mvExamCentreId
          mvActionNumber = 0
          dgr.ContextMenuStrip = mvActionMenu
        Case XMLExamDataSelectionTypes.ExamCentreActionAnalysis
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreActionAnalysis
          vList.IntegerValue("ActionNumber") = mvActionNumber
        Case XMLExamDataSelectionTypes.ExamCentreActionLinks
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreActionLinks
          vList.IntegerValue("ActionNumber") = mvActionNumber
        Case XMLExamDataSelectionTypes.ExamCentreUnits
          vShowGrid = False
        Case XMLExamDataSelectionTypes.ExamExemptions
          vList("SystemColumns") = "N"
          vShowGrid = False
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamExemptions
          vList.IntegerValue("ExamExemptionId") = mvExamExemptionId
          vItemID = mvExamExemptionId
        Case XMLExamDataSelectionTypes.ExamExemptionUnits
          'vShowGrid = False
          vList.IntegerValue("ExamExemptionId") = mvExamExemptionId
        Case XMLExamDataSelectionTypes.ExamSessionCentres
          vShowGrid = False
        Case XMLExamDataSelectionTypes.ExamPersonnel
          vList("SystemColumns") = "N"
          vShowGrid = False
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamPersonnel
          vList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
          vItemID = mvExamPersonnelId
        Case XMLExamDataSelectionTypes.ExamPersonnelAssessmentTypes
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamPersonnelAssessmentTypes
          vList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
        Case XMLExamDataSelectionTypes.ExamPersonnelExpenses
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamPersonnelExpenses
          vList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
        Case XMLExamDataSelectionTypes.ExamPersonnelMarkerInfo
          vList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
          mvExamMaintenanceType = XMLExamMaintenanceTypes.None
          dgr.ContextMenuStrip = mvExamsGridMenu
        Case XMLExamDataSelectionTypes.ExamUnitCandidates
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
          mvExamMaintenanceType = XMLExamMaintenanceTypes.None
        Case XMLExamDataSelectionTypes.ExamSchedule
          If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then
            vShowGrid = False
          Else
            mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamSchedule
          End If
          vList.IntegerValue("ExamSessionId") = mvCourseSessionId
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case XMLExamDataSelectionTypes.ExamSessions
          vList("SystemColumns") = "N"
          vShowGrid = False
          mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamSessions
          vList.IntegerValue("ExamSessionId") = mvExamSessionId
          vItemID = mvExamSessionId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitAssessmentTypes
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitAssessmentTypes
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitEligibilityChecks
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitRequirements
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitGrades
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitGrades
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitPersonnel
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitPersonnel
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitPrerequisites
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitPrerequisites
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitMarkerAllocation
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitMarkerAllocation
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
          vList.Add("GetUnallocatedCount", "Y")
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitResources
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitResources
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnits
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnit
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
          vList.Add("ExamUnitLinkId", mvExamSelectorItem.LinkID)
          vList("SystemColumns") = "N"
          vShowGrid = False
          vItemID = mvExamUnitId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCentreCategories
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreCategories
          vList.IntegerValue("ExamCentreId") = mvExamCentreId
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitLinkCategories
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitLinkCategories
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCentreUnitLinkCategories
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories
          vList.IntegerValue("ExamCentreId") = mvExamCentreId
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
          vList.IntegerValue("ExamCentreUnitId") = mvExamSelectorItem.CentreUnitID
          vList.IntegerValue("ExamUnitLinkId") = mvExamSelectorItem.LinkID
        Case XMLExamDataSelectionTypes.ExamCentreUnitDetails
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreUnitDetails
          vList.IntegerValue("ExamUnitId") = mvExamSelectorItem.UnitID
          vList("SystemColumns") = "N"
          vList.Add("ExamUnitLinkId", mvExamSelectorItem.LinkID)
          vList.Add("ExamCentreUnitId", mvExamSelectorItem.CentreUnitID)
          vShowGrid = False
          'vItemID = mvExamSelectorItem.ID
        Case XMLExamDataSelectionTypes.ExamUnitLinkDocuments,
          XMLExamDataSelectionTypes.ExamCentreDocuments,
          XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments

          mvDocumentMenu = New DocumentMenu(Me)
          dgr.ContextMenuStrip = mvDocumentMenu

          If mvExamDataType = XMLExamDataSelectionTypes.ExamUnitLinkDocuments Then
            vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
            mvDocumentMenu.ExamUnitLinkId = sel.GetLinkID
          ElseIf mvExamDataType = XMLExamDataSelectionTypes.ExamCentreDocuments Then
            vList.IntegerValue("ExamCentreId") = mvExamCentreId
            Dim vExamCentreList As New ParameterList(True)
            vExamCentreList.Add("ExamCentreId", mvExamCentreId)
            Dim vDS As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamCentres, vList)

            If vDS IsNot Nothing AndAlso vDS.Tables.Contains("DataRow") Then
              If vDS.Tables("DataRow")(0).Item("OrganisationNumber").ToString.Length > 0 Then mvDocumentMenu.ExamCentreOrginisation = CInt(vDS.Tables("DataRow")(0).Item("OrganisationNumber").ToString)
              If vDS.Tables("DataRow")(0).Item("ContactNumber").ToString.Length > 0 Then mvDocumentMenu.ExamCentreContact = CInt(vDS.Tables("DataRow")(0).Item("ContactNumber").ToString)
            End If
            mvDocumentMenu.ExamCentreId = mvExamCentreId
          Else
            Dim vExamCentreList As New ParameterList(True)
            vExamCentreList.Add("ExamCentreUnitId", mvExamSelectorItem.CentreUnitID)
            Dim vDS As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamCentreUnits, vExamCentreList)
            If vDS IsNot Nothing AndAlso vDS.Tables.Contains("DataRow") Then
              vExamCentreList = New ParameterList(True)
              vExamCentreList.Add("ExamCentreId", vDS.Tables("DataRow")(0).Item("ExamCentreId").ToString)
              Dim vExCentreDataSet As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamCentres, vExamCentreList)

              If vExCentreDataSet IsNot Nothing AndAlso vExCentreDataSet.Tables.Contains("DataRow") Then
                If vExCentreDataSet.Tables("DataRow")(0).Item("OrganisationNumber").ToString.Length > 0 Then mvDocumentMenu.ExamCentreOrginisation = CInt(vExCentreDataSet.Tables("DataRow")(0).Item("OrganisationNumber").ToString)
                If vExCentreDataSet.Tables("DataRow")(0).Item("ContactNumber").ToString.Length > 0 Then mvDocumentMenu.ExamCentreContact = CInt(vExCentreDataSet.Tables("DataRow")(0).Item("ContactNumber").ToString)
              End If
            End If

            vList.IntegerValue("ExamCentreUnitId") = mvExamSelectorItem.CentreUnitID
            mvDocumentMenu.ExamCentreUnitId = mvExamSelectorItem.CentreUnitID
          End If
          ShowDocumentDetails(0)
          vShowPanel = True
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitStudyModes
          vShowGrid = False
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitStudyModes
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCentreUnitStudyModes
          vShowGrid = False
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitStudyModes
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
          vList.IntegerValue("ExamCentreUnitLinkId") = mvExamSelectorItem.CentreUnitID
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitCertRunTypes
          mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitCertRunTypes
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCustomForm
          'Don't set mvMaintenanceType here, retain what it is so that you know what sort of custom form this is.
          tabMaster.SelectedTab = tabCustomForm
          Dim vDataContext As New CustomFormDataContext(mvExamSelectorItem.CustomFormID, mvExamSelectorItem.ID)
          pnlCustomFormPage.Init(vDataContext)
        Case Else
          Throw New InvalidOperationException("Unexpected data type.")
      End Select
      splRight.Panel1Collapsed = Not vShowGrid And (mvExamDataType <> XMLExamDataSelectionTypes.ExamCentres And
                                                    mvExamDataType <> XMLExamDataSelectionTypes.ExamCentreUnitDetails And
                                                    mvExamDataType <> XMLExamDataSelectionTypes.ExamUnits)
      cmdUnallocate.Visible = False
      cmdAllocate.Visible = False
      cmdLink.Visible = (mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions)
      cmdAnalysis.Visible = (mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions)
      cmdClose.Visible = (mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionLinks Or mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionAnalysis)
      mvCurrentEpl.SuspendLayout()
      Try
        Select Case mvExamDataType
          Case XMLExamDataSelectionTypes.ExamCentreUnits, XMLExamDataSelectionTypes.ExamSessionCentres, XMLExamDataSelectionTypes.ExamExemptionUnits
            mvCurrentEpl.Visible = False
            dgrDetails.Visible = False
            selExams.Visible = True
            cmdSelectAll.Visible = mvExamDataType <> XMLExamDataSelectionTypes.ExamExemptionUnits
            cmdUnSelectAll.Visible = cmdSelectAll.Visible
            cmdNew.Visible = False
            cmdNewChild.Visible = False
            cmdDelete.Visible = False
            cmdSave.Visible = True
            mvSelExamsMenu.SetContext(mvExamDataType, mvExamUnitId, selExams.GetParentID, mvCourseSessionId, True, 0)
            Select Case mvExamDataType
              Case XMLExamDataSelectionTypes.ExamCentreUnits
                selExams.Init(ExamSelector.SelectionType.CentreCourses, mvExamCentreId)
              Case XMLExamDataSelectionTypes.ExamExemptionUnits
                selExams.Init(ExamSelector.SelectionType.ExemptionCourses, mvExamExemptionId)
                cmdDelete.Visible = True
              Case XMLExamDataSelectionTypes.ExamSessionCentres
                selExams.Init(ExamSelector.SelectionType.SessionCentres, mvExamSessionId)
            End Select
          Case XMLExamDataSelectionTypes.ExamPersonnelMarkerInfo, XMLExamDataSelectionTypes.ExamUnitCandidates
            mvCurrentEpl.Visible = False
            dgrDetails.Visible = False
            cmdNew.Visible = False
            cmdDelete.Visible = False
            cmdSave.Visible = False
          Case XMLExamDataSelectionTypes.ExamUnitMarkerAllocation
            mvCurrentEpl.Visible = False
            dgrDetails.Clear()
            dgrDetails.Visible = True
            cmdNew.Visible = False
            cmdDelete.Visible = False
            cmdSave.Visible = False
            cmdNewChild.Visible = False
            cmdSelectAll.Visible = True
            cmdUnSelectAll.Visible = True
            cmdUnallocate.Visible = True
            cmdSelectAll.Enabled = False
            cmdUnSelectAll.Enabled = False
            cmdUnallocate.Enabled = False
            cmdAllocate.Enabled = False
          Case Else
            mvCurrentEpl.Init(New EditPanelInfo(mvExamMaintenanceType))
            mvCurrentEpl.Visible = True
            dgrDetails.Visible = False
            mvCustomiseMenu.SetContext(mvExamMaintenanceType)
            mvCurrentEpl.ContextMenuStrip = mvCustomiseMenu
            selExams.Visible = False
            cmdSelectAll.Visible = False
            cmdUnSelectAll.Visible = False
            cmdNew.Visible = True
            cmdDelete.Visible = True
            cmdNewChild.Visible = False
            cmdSave.Visible = True
        End Select

        'Set the panel up and ensure that the controls that are only relevant to the NG Grading method are hidden if the Concept Grading method's been chosen.
        HideControlsByGradingMethod()

      Finally
        mvCurrentEpl.ResumeLayout()
      End Try

      If mvExamDataType = XMLExamDataSelectionTypes.ExamCentres Then
        Me.dspTabGrid.Visible = True
      ElseIf mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitDetails OrElse
        mvExamDataType = XMLExamDataSelectionTypes.ExamUnits Then
        Me.dgrDetails.Visible = True
      End If
      If mvExamDataType <> XMLExamDataSelectionTypes.ExamExemptionUnits AndAlso mvExamDataType <> XMLExamDataSelectionTypes.ExamPersonnelMarkerInfo AndAlso mvExamDataType <> XMLExamDataSelectionTypes.ExamCentreUnitDetails Then
        vList.IntegerValue("ExamUnitId") = mvExamUnitId
      End If
      Select Case mvExamDataType
        Case XMLExamDataSelectionTypes.ExamCentres
          Dim vTextLookupBox As TextLookupBox = mvCurrentEpl.FindTextLookupBox("ContactNumber")
          vTextLookupBox.SetOrganisationContacts(sel.GetID2)
        Case XMLExamDataSelectionTypes.ExamCentreContacts
          Dim vTextLookupBox As TextLookupBox = mvCurrentEpl.FindTextLookupBox("ContactNumber")
          vTextLookupBox.SetOrganisationContacts(sel.GetParentID2)
        Case XMLExamDataSelectionTypes.ExamUnitGrades
          mvCurrentEpl.FindTextLookupBox("GradeUnits").MultipleValuesSupported = True
        Case XMLExamDataSelectionTypes.ExamUnits, XMLExamDataSelectionTypes.ExamCentreUnitDetails
          mvCurrentEpl.FindPanelControl(Of TextLookupBox)("ActivityGroup", True).FillComboWithRestriction("C")
      End Select
      mvExamsMenu.SetContext(mvExamDataType, mvExamUnitId, sel.GetParentID, mvCourseSessionId, False, vItemID)
      mvExamsGridMenu.SetContext(mvExamDataType, mvExamUnitId, sel.GetParentID, mvCourseSessionId, False, vItemID)
      If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then
        cmdNew.Visible = False
        cmdDelete.Visible = False
        cmdSelectAll.Visible = True
        cmdUnSelectAll.Visible = True
        mvSelectedRow = -1
        SetDefaults()
        mvExamScheduleSelector = DirectCast(FindControl(mvCurrentEpl, "ExamCentreCode"), ExamSelector)
        mvExamScheduleSelector.Init(ExamSelector.SelectionType.ScheduleCentres, mvCourseSessionId, mvExamUnitId)
        'ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitStudyModes Then
      ElseIf mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitStudyModes Or
             mvExamDataType = XMLExamDataSelectionTypes.ExamUnitStudyModes Then
        vDataSet = ExamsDataHelper.GetExamData(mvExamDataType, vList)
        Dim vStudyModeList As CheckedListBox = DirectCast(mvCurrentEpl.Controls("UnitStudyModes"), CheckedListBox)
        Dim vStudyModesNotAllowed As TransparentLabel = DirectCast(mvCurrentEpl.Controls("StudyModesNotAllowed_Label"), TransparentLabel)
        Dim vStudyModesWarning As TransparentLabel = DirectCast(mvCurrentEpl.Controls("StudyModesWarning_Label"), TransparentLabel)
        vStudyModeList.Items.Clear()
        mvStudyModeList = New List(Of String)
        If vDataSet.Tables.Contains("DataRow") Then
          For Each vRow As DataRow In vDataSet.Tables("DataRow").Rows
            vStudyModeList.Items.Add(vRow("StudyModeDesc"), CStr(vRow("Selected")) = "Y")
            mvStudyModeList.Add(CStr(vRow("StudyMode")))
          Next vRow
        End If
        If vStudyModeList.Items.Count > 0 Then
          cmdSave.Visible = True
          vStudyModeList.Visible = True
          vStudyModesNotAllowed.Visible = False
        Else
          cmdSave.Visible = False
          vStudyModeList.Visible = False
          vStudyModesNotAllowed.Visible = True
        End If
        vStudyModesWarning.Visible = mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitStudyModes And
                                     vStudyModeList.Visible And
                                     vStudyModeList.CheckedItems.Count < 1
        AddHandler vStudyModeList.ItemCheck, Sub(sender As Object, e As ItemCheckEventArgs)
                                               vStudyModesWarning.Visible = mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitStudyModes And
                                                                            vStudyModeList.Visible And
                                                                            vStudyModeList.CheckedItems.Count + If(e.NewValue = CheckState.Checked, 1, -1) < 1
                                             End Sub
        cmdNew.Visible = False
        cmdDelete.Visible = False
      Else
        vDataSet = ExamsDataHelper.GetExamData(mvExamDataType, vList)
        If vShowGrid Then dgr.Populate(vDataSet)
        Select Case mvExamDataType
          Case XMLExamDataSelectionTypes.ExamCentreDocuments, XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments, XMLExamDataSelectionTypes.ExamUnitLinkDocuments
            If mvExamDataType = XMLExamDataSelectionTypes.ExamCentreDocuments Then mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreDocuments
            If mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments Then mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreUnitLinkDocuments
            If mvExamDataType = XMLExamDataSelectionTypes.ExamUnitLinkDocuments Then mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitLinkDocuments
            mvDocumentsDataSet = vDataSet

          Case XMLExamDataSelectionTypes.ExamCentreUnitLinkCategories
            SetRowReadOnly()

          Case XMLExamDataSelectionTypes.ExamUnitStudyModes
            mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitStudyModes
        End Select
        If vShowPanel Then
          dpl.Visible = True
          splTab.Visible = True
          dpl.Init(vDataSet, False, False)
        End If
        If vDataSet.Tables.Contains("DataRow") Then
          If vShowGrid Then
            SelectRow(0)
          Else
            If Not mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnits Then
              mvCurrentEpl.Populate(vDataSet.Tables("DataRow").Rows(0))
              mvSelectedRow = 0
            End If
          End If
          If Not mvExamDataType = XMLExamDataSelectionTypes.ExamCentreUnitLinkCategories Then cmdDelete.Enabled = True
        Else
          cmdDelete.Enabled = False
          mvSelectedRow = -1
          SetDefaults()
        End If
      End If
      Select Case mvExamDataType
        Case XMLExamDataSelectionTypes.ExamUnits
          If mvExamUnitId > 0 Then
            mvCurrentEpl.EnableControlList("ExamUnitCode,ExamUnitType", False)
            If mvCurrentEpl.FindTextLookupBox("ExamUnitType").GetDataRowItem("ExamQuestion").ToString = "Y" Then
              mvCurrentEpl.EnableControl("ExamMarkType", False)
              cmdNewChild.Visible = False
            Else
              mvCurrentEpl.EnableControl("ExamMarkType", True)
              cmdNewChild.Visible = True
            End If
            Dim vTempParms As New ParameterList(True, True)
            vTempParms.Add("ExamUnitLinkId", mvExamSelectorItem.LinkID)
            dgrDetails.Populate(ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamAccreditationHistory, vTempParms))
          End If
        Case XMLExamDataSelectionTypes.ExamCentres
          If mvExamCentreId > 0 Then
            cmdNewChild.Visible = True
            mvCurrentEpl.EnableControlList("ExamCentreCode", False)
            mvCentreOrganisation = IntegerValue(mvCurrentEpl.GetValue("OrganisationNumber"))
            mvCentreContact = IntegerValue(mvCurrentEpl.GetValue("ContactNumber"))
            Dim vTempParms As New ParameterList(True, True)
            vTempParms.Add("ExamCentreId", mvExamCentreId.ToString)
            dspTabGrid.SetText("")
            dspTabGrid.DisplayGrid(0).Populate(ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamCentreHistory, vTempParms))
            Dim vAccHistory As New ParameterList(True, True)
            vAccHistory.Add("ExamCentreId", mvExamCentreId)
            dspTabGrid.DisplayGrid(1).Populate(ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamAccreditationHistory, vAccHistory))
            dspTabGrid.SetTabPages(3)
          End If
        Case XMLExamDataSelectionTypes.ExamSessions
          If mvExamSessionId > 0 Then
            mvCurrentEpl.EnableControlList("ExamSessionCode", False)
          End If
        Case XMLExamDataSelectionTypes.ExamExemptions
          If mvExamExemptionId > 0 Then
            mvCurrentEpl.EnableControlList("ExamExemptionCode", False)
          End If
        Case XMLExamDataSelectionTypes.ExamCentreUnitDetails
          mvCurrentEpl.EnableControls(False)
          mvCurrentEpl.EnableControlList("LocalName,CourseAccreditation,CourseAccreditationValidFrom,CourseAccreditationValidTo", True)
          cmdNew.Visible = False
          cmdNewChild.Visible = False
          cmdDelete.Visible = False
          Dim vTempParms As New ParameterList(True, True)
          vTempParms.Add("ExamCentreUnitId", mvExamSelectorItem.CentreUnitID)
          dgrDetails.Populate(ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamAccreditationHistory, vTempParms))
        Case XMLExamDataSelectionTypes.ExamUnitLinkDocuments, XMLExamDataSelectionTypes.ExamCentreDocuments, XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments
          dspTabGrid.Visible = True
          dgr.SetToolBarVisible()
          cmdDelete.Visible = False
          cmdSave.Visible = False
          cmdNew.Visible = False
          Select Case mvExamDataType
            Case XMLExamDataSelectionTypes.ExamUnitLinkDocuments
              mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamUnitLinkDocuments
            Case XMLExamDataSelectionTypes.ExamCentreDocuments
              mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreDocuments
            Case Else
              mvExamMaintenanceType = ExamsAccess.XMLExamMaintenanceTypes.ExamCentreUnitLinkDocuments
          End Select
        Case XMLExamDataSelectionTypes.ExamCentreActionLinks
          SetDefaults()
      End Select
      cmdLink.Visible = (mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions)
      cmdAnalysis.Visible = (mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions)
      cmdLink.Enabled = (mvActionNumber <> 0)
      cmdAnalysis.Enabled = (mvActionNumber <> 0)
      If mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions Or
        mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionLinks Or
        mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionAnalysis Then
        cmdDelete.Enabled = (dgr.DataRowCount > 0)
      End If
      bpl.RepositionButtons()
      ResizePanels()
    Finally
      Me.splTop.Panel2.ResumeLayout()
      Me.splTop.ResumeLayout()
      Me.splBottom.Panel1.ResumeLayout()
      Me.splBottom.Panel2.ResumeLayout()
      Me.splBottom.ResumeLayout()
      Me.pnlCommands.ResumeLayout()
      Me.tabMaster.ResumeLayout()
      Me.tabMain.ResumeLayout()
      Me.splRight.Panel1.ResumeLayout()
      Me.splRight.Panel2.ResumeLayout()
      Me.splRight.ResumeLayout()
      Me.bpl.ResumeLayout()
      Me.tabCustomForm.ResumeLayout()
      Me.dgrMenuStrip.ResumeLayout()
      Me.dgr2MenuStrip.ResumeLayout()
      Me.dplMenuStrip.ResumeLayout()
      Me.dgr1MenuStrip.ResumeLayout()
      Me.dgr0MenuStrip.ResumeLayout()
      Me.ResumeLayout()
    End Try
  End Sub

  Private Sub ShowDocumentDetails(ByVal pNumber As Integer)
    If pNumber > 0 Then
      mvDocumentMenu.DocumentNumber = pNumber
      dspTabGrid.TabPageText(0) = ControlText.TbpPrecis
      dspTabGrid.SetText(DataHelper.GetDocumentPrecis(pNumber))

      mvDocAnalysisDataSource = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentSubjects, pNumber)
      dspTabGrid.DisplayGrid(0).Populate(mvDocAnalysisDataSource)
      dspTabGrid.GridContextMenuStrip(0) = dgr0MenuStrip
      dspTabGrid.DisplayGrid(0).SetToolBarVisible()

      mvDocLinkDataSource = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentLinks, pNumber)
      dspTabGrid.DisplayGrid(1).Populate(mvDocLinkDataSource)
      dspTabGrid.GridContextMenuStrip(1) = dgr1MenuStrip
      dspTabGrid.DisplayGrid(1).SetToolBarVisible()

      mvDocumentMenu.SetNotifyProcessed(dspTabGrid.DisplayGrid(1))
      dspTabGrid.DisplayGrid(2).Populate(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentHistory, pNumber))
      dspTabGrid.DisplayGrid(2).MaxGridRows = DisplayTheme.HistoryMaxGridRows
      dspTabGrid.DisplayGrid(2).SetToolBarVisible()
      dspTabGrid.SetTabPages(4)
    Else
      dspTabGrid.TabPageText(0) = String.Empty
      dspTabGrid.DisplayGrid(0).Clear()
      dspTabGrid.DisplayGrid(1).Clear()
      dspTabGrid.DisplayGrid(2).MaxGridRows = DisplayTheme.HistoryMaxGridRows
      dspTabGrid.DisplayGrid(2).Clear()
    End If
    dspTabGrid.SetTabPages(4)
  End Sub

  Private Sub mvDocumentMenu_ShowRelatedDocument(ByVal Sender As Object) Handles mvDocumentMenu.ShowRelatedDocument
    If mvDocumentMenu.DocumentNumber > 0 Then
      Dim vList As New ParameterList
      vList.IntegerValue("CommunicationsLogNumber1") = mvDocumentMenu.DocumentNumber
      vList("FinderCaption") = ControlText.FrmRelatedDocumentsFinder
      FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftDocuments, vList)
    End If
  End Sub

  Private Sub PopulateSessions()
    cboSessions.ValueMember = "ExamSessionId"
    cboSessions.DisplayMember = "ExamSessionDescription"
    Dim vList As New ParameterList(True)
    DataHelper.FillComboBox(cboSessions, CareNetServices.XMLLookupDataTypes.xldtExamSessions, True, vList)
    If cboSessions.DataSource IsNot Nothing Then
      Dim vRow As DataRow = DirectCast(cboSessions.DataSource, DataTable).Rows(0)
      vRow("ExamSessionDescription") = "<No Session>"
    End If
  End Sub

#End Region

#Region "Edit Panel Events"

  Private Sub epl_ContactSelected(ByVal pSender As Object, ByVal pContactNumber As Integer) Handles mvCurrentEpl.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub epl_GetInitialCodeRestrictions(ByVal pSender As Object, ByVal pParameterName As String, ByRef pList As ParameterList) Handles mvCurrentEpl.GetInitialCodeRestrictions
    Select Case pParameterName
      Case "ExamCentreCode"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
        If mvCourseSessionId > 0 Then
          pList.IntegerValue("ExamSessionId") = mvCourseSessionId
        End If
      Case "GeographicalRegion"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList("GeographicalRegionType") = AppValues.ControlValue(AppValues.ControlTables.exam_controls, AppValues.ControlValues.geographical_region_type)
      Case "GradeUnits"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList("AllUnits") = "Y"
      Case "ExamPrerequisiteUnitCode"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList("AllUnits") = "Y"
      Case "StandardDocument"
        If Me.mvExamDataType = XMLExamDataSelectionTypes.ExamUnitCertRunTypes Then
          If pList Is Nothing Then pList = New ParameterList(True)
          pList("ExamStandardDocument") = "Y"
        End If
    End Select
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal pSender As Object, ByVal pParameterName As String, ByVal pList As ParameterList) Handles mvCurrentEpl.GetCodeRestrictions
    Select Case pParameterName
      Case "Product"
        pList("FindProductType") = "Q"      'Exam
      Case "ExamPersonnelId"
        If mvExamDataType = XMLExamDataSelectionTypes.ExamUnitPersonnel AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ex_restrict_personnel) Then
          Dim vList As New ParameterList(True)
          vList.IntegerValue("ExamUnitId") = mvExamUnitId
          Dim vTable As DataTable = ExamsDataHelper.SelectExamData(XMLExamDataSelectionTypes.ExamUnitAssessmentTypes, vList)
          If vTable IsNot Nothing Then
            Dim vAssessTypes As New ArrayListEx
            For Each vRow As DataRow In vTable.Rows
              vAssessTypes.Add(vRow("ExamAssessmentType").ToString)
            Next
            If vAssessTypes.Count > 0 Then
              pList("ExamAssessmentType") = vAssessTypes.CSList
            End If
          End If
        End If
    End Select
  End Sub

  Private Sub epl_ValueChange(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles mvCurrentEpl.ValueChanged
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamCentres
        Select Case pParameterName
          Case "OrganisationNumber"
            mvCurrentEpl.FindTextLookupBox("ContactNumber").SetOrganisationContacts(IntegerValue(mvCurrentEpl.GetValue("OrganisationNumber")))
          Case "AddressNumber"
            If Not EditingExistingRecord() Then
              Dim vComboBox As ComboBox = mvCurrentEpl.FindComboBox("AddressNumber")
              If vComboBox.SelectedItem IsNot Nothing Then
                If DirectCast(vComboBox.SelectedItem, DataRowView).Item("UK").ToString.StartsWith("Y") Then
                  mvCurrentEpl.SetValue("Overseas", "N")
                Else
                  mvCurrentEpl.SetValue("Overseas", "Y")
                End If
              End If
            End If
          Case "AccreditationStatus", "CourseAccreditation"
            If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidFrom") AndAlso Not mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked Then
              mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked = True
            End If
        End Select
      Case XMLExamMaintenanceTypes.ExamUnit
        Select Case pParameterName
          Case "ExamUnitType"
            Dim vUnitTypeTextLookupBox As TextLookupBox = mvCurrentEpl.FindTextLookupBox("ExamUnitType")
            If vUnitTypeTextLookupBox.GetDataRowItem("ExamQuestion").ToString = "Y" Then
              mvCurrentEpl.SetValue("ExamMarkType", "M", True)
              mvCurrentEpl.EnableControl("MarkFactor", True)
            Else
              mvCurrentEpl.EnableControl("ExamMarkType", True)
            End If
            mvCurrentEpl.PanelInfo.PanelItems("Product").Mandatory = vUnitTypeTextLookupBox.GetDataRowItem("ScheduleRequired") = "Y"
            mvCurrentEpl.PanelInfo.PanelItems("Rate").Mandatory = vUnitTypeTextLookupBox.GetDataRowItem("ScheduleRequired") = "Y"
            mvCurrentEpl.SetErrorField(pParameterName, "") 'Note that the error needs to be cleared before validation is fired up.  See the ValidateItem event for ExamUnitType validation
          Case "ExamMarkType"
            mvCurrentEpl.EnableControl("MarkFactor", pValue = "M")
          Case "ExamMarkerStatus"
            mvCurrentEpl.PanelInfo.PanelItems("PapersPerMarker").Mandatory = (pValue <> "N" AndAlso pValue <> "E")
          Case "TimeLimitType"
            mvCurrentEpl.PanelInfo.PanelItems("UnitTimeLimit").Mandatory = (pValue = "M" Or pValue = "Y")
          Case "UnitTimeLimit"
            mvCurrentEpl.PanelInfo.PanelItems("TimeLimitType").Mandatory = pValue.Length > 0
          Case "AccreditationStatus", "CourseAccreditation"
            If pValue.Length = 0 AndAlso mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidFrom") AndAlso mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked Then mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked = False
            If pValue.Length = 0 AndAlso mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidTo") AndAlso mvCurrentEpl.FindDateTimePicker("AccreditationValidTo").Checked Then mvCurrentEpl.FindDateTimePicker("AccreditationValidTo").Checked = False

            If pValue.Length > 0 AndAlso mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidFrom") AndAlso Not mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked Then
              mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked = True
            End If
        End Select
      Case XMLExamMaintenanceTypes.ExamUnitPersonnel
        If pParameterName = "ExamPersonnelId" Then
          Dim vTextLookupBox As TextLookupBox = mvCurrentEpl.FindTextLookupBox(pParameterName)
          mvCurrentEpl.SetValue("ExamPersonnelType", vTextLookupBox.GetDataRowItem("ExamPersonnelType"))
        End If
      Case XMLExamMaintenanceTypes.ExamUnitGrades
        Select Case pParameterName
          Case "ExamGrade"
            Dim vRow As Integer = dgr.FindRow(pParameterName, pValue)
            If vRow >= 0 Then
              mvCurrentEpl.SetValue("SequenceNumber", dgr.GetValue(vRow, "SequenceNumber"))
            Else
              Dim vTextLookupBox As TextLookupBox = mvCurrentEpl.FindTextLookupBox(pParameterName)
              mvCurrentEpl.SetValue("SequenceNumber", vTextLookupBox.GetDataRowItem("SequenceNumber"))
            End If
          Case "ExamGradeConditionType"
            SetGradingRuleRequiredUI(pValue, False)
        End Select
      Case XMLExamMaintenanceTypes.ExamCentreUnitDetails
        If pValue.Length = 0 AndAlso mvCurrentEpl.PanelInfo.PanelItems.Exists("CourseAccreditationValidFrom") AndAlso mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidFrom").Checked Then mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidFrom").Checked = False
        If pValue.Length = 0 AndAlso mvCurrentEpl.PanelInfo.PanelItems.Exists("CourseAccreditationValidTo") AndAlso mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidTo").Checked Then mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidTo").Checked = False

        If pParameterName = "CourseAccreditation" Then
          If pValue.Length > 0 AndAlso mvCurrentEpl.PanelInfo.PanelItems.Exists("CourseAccreditationValidFrom") AndAlso Not mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidFrom").Checked Then
            mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidFrom").Checked = True
          End If
        End If

    End Select
  End Sub

  Private Sub SetGradingRuleRequiredUI(ByVal pValue As String, ByVal pSetValue As Boolean)
    Dim vIsGrade As Boolean = pValue = "G"
    Dim vIsResult As Boolean = pValue = "P"
    Dim vIsStatus As Boolean = pValue = "S"
    Dim vIsMark As Boolean = pValue = "M"
    Dim vIsOther As Boolean = Not (vIsGrade Or vIsResult Or vIsStatus Or vIsMark)

    If pSetValue Then
      If vIsGrade Then
        mvCurrentEpl.SetValue("RequiredGrade", mvCurrentEpl.GetValue("RequiredValue"))
      ElseIf vIsResult Then
        mvCurrentEpl.SetValue("RequiredResult", mvCurrentEpl.GetValue("RequiredValue"))
      ElseIf vIsStatus Then
        mvCurrentEpl.SetValue("RequiredStatus", mvCurrentEpl.GetValue("RequiredValue"))
      ElseIf vIsMark Then
        mvCurrentEpl.SetValue("RequiredMark", mvCurrentEpl.GetValue("RequiredValue"))
      End If
    End If
    mvCurrentEpl.SetControlVisible("RequiredGrade", vIsGrade)
    mvCurrentEpl.PanelInfo.PanelItems.SetOptionalItemMandatory("RequiredGrade", vIsGrade)
    mvCurrentEpl.SetControlVisible("RequiredResult", vIsResult)
    mvCurrentEpl.PanelInfo.PanelItems.SetOptionalItemMandatory("RequiredResult", vIsResult)
    mvCurrentEpl.SetControlVisible("RequiredStatus", vIsStatus)
    mvCurrentEpl.PanelInfo.PanelItems.SetOptionalItemMandatory("RequiredStatus", vIsStatus)
    mvCurrentEpl.SetControlVisible("RequiredMark", vIsMark)
    mvCurrentEpl.PanelInfo.PanelItems.SetOptionalItemMandatory("RequiredMark", vIsMark)
    mvCurrentEpl.SetControlVisible("RequiredValue", vIsOther)
    mvCurrentEpl.PanelInfo.PanelItems.SetOptionalItemMandatory("RequiredValue", vIsOther)
    Dim vTextBox As TextBox = mvCurrentEpl.FindTextBox("RequiredValue")
    If vIsOther Or vIsMark Then
      Dim vResult As Double
      If Double.TryParse(mvCurrentEpl.GetValue("RequiredValue"), vResult) = False Then mvCurrentEpl.SetValue("RequiredValue", "")
      AddHandler vTextBox.KeyPress, AddressOf NumericKeyPressHandler
    Else
      RemoveHandler vTextBox.KeyPress, AddressOf NumericKeyPressHandler
    End If
  End Sub

  Private Sub ExamScheduleSelected(ByVal sender As Object, ByVal pType As ExamsAccess.XMLExamDataSelectionTypes, ByVal pItem As ExamSelectorItem) Handles mvExamScheduleSelector.ItemSelected
    Try
      If mvExamScheduleSelector.GetID2 > 0 Then
        Dim vList As New ParameterList(True)
        vList.IntegerValue("ExamScheduleId") = mvExamScheduleSelector.GetID2
        Dim vRow As DataRow = DataHelper.GetRowFromDataSet(ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamSchedule, vList))
        If vRow IsNot Nothing Then mvCurrentEpl.Populate(vRow)
      End If
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

#End Region

#Region "Grid Handling"

  Private Sub dgr_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrDetails_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgrDetails.ContactSelected
    Try
      FormHelper.ShowContactCardIndex(pContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrDetails_CheckBoxClicked(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String) Handles dgrDetails.CheckBoxClicked
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamUnitMarkerAllocation
        If dgrDetails.GetColumn(mvSelectColumnName) = pCol Then
          If dgrDetails.GetValue(pRow, pCol) = "True" Then
            dgrDetails.SetValue(pRow, pCol, "")
          Else
            dgrDetails.SetValue(pRow, pCol, "True")
          End If
        End If
    End Select
  End Sub

  Private Sub dgr_CanCustomise(ByVal pSender As Object, ByVal pRow As String) Handles dgr.CanCustomise
    RefreshCard()
    SetNodeDataEditable(sel.SelectedNode)
  End Sub

  Private Sub dgr_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    'If anything changed then confirm cancel changes
    SelectRow(pRow)
  End Sub

  Private Sub dgr_RowChanging(ByVal pSender As Object, ByRef pCancel As Boolean) Handles dgr.RowChanging
    If mvCurrentEpl.DataChanged Then
      If ConfirmSave() Then
        pCancel = Not ProcessSave(pSender)
      End If
    End If
  End Sub

  Private Sub SelectRow(ByVal pRow As Integer)
    Try
      dpl.SuspendLayout()
      If pRow >= 0 Then
        Dim vList As New ParameterList(True)
        Dim vCount As Integer = vList.Count
        GetPrimaryKeyValues(vList, pRow, False)
        If vList.Count > vCount Then
          Select Case mvExamMaintenanceType
            Case XMLExamMaintenanceTypes.ExamCentreAssessmentTypes,
                 XMLExamMaintenanceTypes.ExamCentreActions,
                 XMLExamMaintenanceTypes.ExamCentreActionLinks,
                 XMLExamMaintenanceTypes.ExamCentreActionAnalysis,
                 XMLExamMaintenanceTypes.ExamCentreContacts,
                 XMLExamMaintenanceTypes.ExamExemptions,
                 XMLExamMaintenanceTypes.ExamPersonnelAssessmentTypes,
                 XMLExamMaintenanceTypes.ExamPersonnelExpenses,
                 XMLExamMaintenanceTypes.ExamSchedule,
                 XMLExamMaintenanceTypes.ExamUnitAssessmentTypes,
                 XMLExamMaintenanceTypes.ExamUnitGrades,
                 XMLExamMaintenanceTypes.ExamUnitPersonnel,
                 XMLExamMaintenanceTypes.ExamUnitPrerequisites,
                 XMLExamMaintenanceTypes.ExamUnitRequirements,
                 XMLExamMaintenanceTypes.ExamUnitResources,
                 XMLExamMaintenanceTypes.ExamUnitMarkerAllocation,
                 XMLExamMaintenanceTypes.ExamUnitLinkCategories,
                 XMLExamMaintenanceTypes.ExamCentreCategories,
                 XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories,
                 XMLExamMaintenanceTypes.ExamUnitLinkDocuments,
                 XMLExamMaintenanceTypes.ExamCentreDocuments,
                 XMLExamMaintenanceTypes.ExamCentreUnitLinkDocuments,
                 XMLExamMaintenanceTypes.ExamUnitCertRunTypes
              If mvExamDataType <> XMLExamDataSelectionTypes.ExamExemptionUnits Then
                Dim vDataTable As DataTable
                Select Case mvExamMaintenanceType
                  Case XMLExamMaintenanceTypes.ExamUnitMarkerAllocation
                    vList.AddSystemColumns()
                    Dim vDataSet As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, vList)
                    dgrDetails.Populate(vDataSet)
                    If dgrDetails.RowCount > 0 Then dgrDetails.SetCheckBoxColumn("Select")
                    cmdAllocate.Visible = (dgrDetails.RowCount > 0 AndAlso dgr.GetValue(pRow, "ContactNumber").Length = 0)
                    cmdUnallocate.Visible = (dgrDetails.RowCount > 0 AndAlso dgr.GetValue(pRow, "ContactNumber").Length > 0)
                    cmdSelectAll.Enabled = (cmdUnallocate.Visible OrElse cmdAllocate.Visible)
                    cmdUnSelectAll.Enabled = (cmdUnallocate.Visible OrElse cmdAllocate.Visible)
                    cmdUnallocate.Enabled = cmdUnallocate.Visible
                    cmdAllocate.Enabled = cmdAllocate.Visible
                  Case XMLExamMaintenanceTypes.ExamUnitLinkCategories, XMLExamMaintenanceTypes.ExamCentreCategories, XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories
                    Dim vExamUnitLink As Integer = 0
                    If vList.ContainsKey("ExamUnitLinkId") Then
                      vExamUnitLink = vList.IntegerValue("ExamUnitLinkId")
                      'vList.Remove("ExamUnitLinkId")
                    End If
                    vDataTable = ExamsDataHelper.SelectExamData(mvExamDataType, vList)
                    mvCurrentEpl.Populate(vDataTable.Rows(0))
                  Case XMLExamMaintenanceTypes.ExamUnitLinkDocuments, XMLExamMaintenanceTypes.ExamCentreDocuments, XMLExamMaintenanceTypes.ExamCentreUnitLinkDocuments
                    If dgr.GetValue(pRow, "DocumentNumber").Length > 0 Then ShowDocumentDetails(CInt(dgr.GetValue(pRow, "DocumentNumber")))
                    If mvDocumentsDataSet IsNot Nothing Then dpl.Populate(mvDocumentsDataSet, dgr.CurrentDataRow)

                  Case XMLExamMaintenanceTypes.ExamCentreActionLinks
                    If dgr.GetValue(pRow, "EntityType").Equals("F", StringComparison.CurrentCultureIgnoreCase) Then
                      'Fundraising Request links cannot be deleted
                      cmdDelete.Enabled = False
                    Else
                      cmdDelete.Enabled = True
                    End If
                  Case Else
                    vDataTable = ExamsDataHelper.SelectExamData(mvExamDataType, vList)
                    If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count = 1 Then
                      If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitPersonnel Then
                        mvCurrentEpl.FindTextLookupBox("ExamPersonnelId").ActiveOnly = False
                      End If
                      mvCurrentEpl.Populate(vDataTable.Rows(0))
                      cmdDelete.Enabled = True
                      Select Case mvExamMaintenanceType
                        Case XMLExamMaintenanceTypes.ExamCentreActions
                          If mvCurrentEpl.GetDoubleValue("DurationDays") = 0 Then
                            mvCurrentEpl.SetValue("DurationDays", String.Empty)
                          End If
                          If mvCurrentEpl.GetDoubleValue("DurationHours") = 0 Then
                            mvCurrentEpl.SetValue("DurationHours", String.Empty)
                          End If
                          mvCurrentEpl.DataChanged = False
                          If Not Integer.TryParse(dgr.GetValue(pRow, "ActionNumber"), mvActionNumber) Then
                            mvActionNumber = 0
                          End If
                          SetActionChangeReason(mvCurrentEpl, (mvActionNumber <> 0), False)
                          If mvActionMenu IsNot Nothing Then
                            mvActionMenu.ActionStatus = dgr.GetValue(pRow, "ActionStatus")
                            If dgr.GetValue(pRow, "ActionNumber").Length > 0 Then mvActionMenu.ActionNumber = CInt(dgr.GetValue(pRow, "ActionNumber"))
                            If dgr.GetValue(pRow, "MasterAction").Length > 0 Then mvActionMenu.MasterActionNumber = CInt(dgr.GetValue(pRow, "MasterAction"))
                          End If
                          cmdDelete.Enabled = mvActionNumber <> 0
                          cmdLink.Enabled = mvActionNumber <> 0
                          cmdAnalysis.Enabled = mvActionNumber <> 0
                        Case XMLExamMaintenanceTypes.ExamUnitGrades
                          SetGradingRuleRequiredUI(mvCurrentEpl.GetValue("ExamGradeConditionType"), True)
                          mvCurrentEpl.DataChanged = False
                        Case XMLExamMaintenanceTypes.ExamUnitPersonnel
                          Dim vTextLookupBox As TextLookupBox = mvCurrentEpl.FindTextLookupBox("ExamPersonnelId")
                          mvCurrentEpl.SetValue("ExamPersonnelType", vTextLookupBox.GetDataRowItem("ExamPersonnelType"), True)
                          mvCurrentEpl.DataChanged = False
                        Case XMLExamMaintenanceTypes.ExamUnitPrerequisites
                          mvCurrentEpl.SetDependancies("PassRequired")
                        Case XMLExamMaintenanceTypes.ExamUnitCertRunTypes 'BR20437
                          mvCurrentEpl.EnableControl("ExamCertRunType", False)
                      End Select
                    Else
                      If dgr.GetValue(pRow, 0).Length > 0 Then
                        'Don't give an error if there was no data on the grid row
                        Throw New CareException("Exams data item not found")
                      End If
                    End If
                End Select
                SetNodeDataEditable(sel.SelectedNode)
              End If

          End Select
        End If
        mvSelectedRow = pRow
      End If
      dpl.ResumeLayout()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_ExamCentreSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCentreId As Integer) Handles dgr.ExamCentreSelected
    Try
      FormHelper.ShowExamIndex(pCentreId, "N")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_ExamUnitSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pUnitlinkId As Integer) Handles dgr.ExamUnitSelected
    Try
      FormHelper.ShowExamIndex(pUnitlinkId, "U")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgr_ExamCentreUnitSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pCentreUnitId As Integer) Handles dgr.ExamCentreUnitSelected
    Try
      FormHelper.ShowExamIndex(pCentreUnitId, "X")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "Maintenance Methods"

  Private Function ProcessSave(ByVal sender As Object) As Boolean
    Dim vChangedItem As New ChangedItem(0, 0, False, "")
    Try
      mvRefreshSelector = False
      Dim vList As New ParameterList(True)
      Select Case mvExamDataType
        Case XMLExamDataSelectionTypes.ExamUnitCertRunTypes
          mvCurrentEpl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
          Try
            If EditingExistingRecord() Then
              ExamsDataHelper.UpdateItem(XMLExamMaintenanceTypes.ExamUnitCertRunTypes, vList)
            Else
              ExamsDataHelper.AddItem(XMLExamMaintenanceTypes.ExamUnitCertRunTypes, vList)
            End If
            mvCurrentEpl.DataChanged = False
          Catch vEx As CareException
            If vEx.ErrorNumber = 10450 Then
              epl.SetErrorField("ExamCertRunType", vEx.Message)
              Return False
            Else
              Throw
            End If
          End Try
          Return True
        Case XMLExamDataSelectionTypes.ExamUnitStudyModes
          Dim vStyudyModeList As CheckedListBox = DirectCast(mvCurrentEpl.Controls("UnitStudyModes"), CheckedListBox)
          Dim vSelectedModes As New List(Of String)
          For vIndex As Integer = 0 To vStyudyModeList.Items.Count - 1
            If vStyudyModeList.GetItemChecked(vIndex) Then
              vSelectedModes.Add(mvStudyModeList(vIndex))
            End If
          Next vIndex
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
          vList("StudyModes") = vSelectedModes.AsCommaSeperated
          Try
            ExamsDataHelper.UpdateItem(ExamsAccess.XMLExamMaintenanceTypes.ExamUnitStudyModes, vList)
            mvCurrentEpl.DataChanged = False
          Catch vEx As CareException
            If vEx.ErrorNumber = CareException.ErrorNumbers.enUnitStudyModeInUse Then
              ShowInformationMessage(vEx.Message)
              mvCurrentEpl.DataChanged = False
              RefreshCard()
              Return False
            Else
              Throw
            End If
          End Try
          Return True
        Case XMLExamDataSelectionTypes.ExamCentreUnitStudyModes
          Dim vStyudyModeList As CheckedListBox = DirectCast(mvCurrentEpl.Controls("UnitStudyModes"), CheckedListBox)
          Dim vSelectedModes As New List(Of String)
          For vIndex As Integer = 0 To vStyudyModeList.Items.Count - 1
            If vStyudyModeList.GetItemChecked(vIndex) Then
              vSelectedModes.Add(mvStudyModeList(vIndex))
            End If
          Next vIndex
          vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
          vList.IntegerValue("ExamCentreUnitLinkId") = mvExamSelectorItem.CentreUnitID
          vList("StudyModes") = vSelectedModes.AsCommaSeperated
          ExamsDataHelper.UpdateItem(ExamsAccess.XMLExamMaintenanceTypes.ExamUnitStudyModes, vList)
          mvCurrentEpl.DataChanged = False
          Return True
        Case XMLExamDataSelectionTypes.ExamCentreActions
          Dim vEditing As Boolean = EditingExistingRecord()
          Dim vValid As Boolean = mvCurrentEpl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)
          If vValid Then
            vList.Add("ExamCentreId", Me.sel.GetParentID)
            If Not vList.Contains("ActionNumber") AndAlso mvActionNumber > 0 Then
              vList("ActionNumber") = mvActionNumber.ToString
            End If
            If FindControl(mvCurrentEpl, "ScheduledOn", False) IsNot Nothing AndAlso mvCurrentEpl.GetValue("ScheduledOn").Length > 0 Then
              'Validate ScheduledOn date for calendar conflicts
              vValid = mvCurrentEpl.GetScheduleDate(vList, CareNetServices.XMLActionScheduleTypes.xastGivenDate, (Not vEditing))
            End If
            vList("Notified") = "N"
            vEditing = mvActionNumber > 0
            Dim mvReturnList = If(vEditing, DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctAction, vList),
                                              DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctAction, vList))
            If mvReturnList.Contains("ActionNumber") Then
              UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetActions, IntegerValue(mvReturnList.Item("ActionNumber")), vList.Item("ActionDesc").ToString)
              Dim mvActionNumber = mvReturnList.IntegerValue("ActionNumber")
            End If
            If Not vEditing AndAlso
               DataHelper.UserInfo.ContactNumber > 0 And mvReturnList("ActionNumber").ToString.Length > 0 Then
              If ShowQuestion(QuestionMessages.QmNoActioners, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                Try
                  'Re-Initialise parameter list to clear all the previous values
                  vList = New ParameterList(True)
                  vList("ActionNumber") = mvReturnList("ActionNumber").ToString
                  vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
                  vList("ActionLinkType") = "A"
                  vList("Notified") = "N"
                  DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
                Catch vEx As CareException
                  If vEx.ErrorNumber = CareException.ErrorNumbers.enAppointmentConflict Then
                    ShowInformationMessage(InformationMessages.ImActionScheduleConflict)
                  Else
                    Throw
                  End If
                End Try
              End If
            End If
            RefreshCard()
          End If
          Return vValid
        Case XMLExamDataSelectionTypes.ExamCentreActionAnalysis
          If mvCurrentEpl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtNone) Then
            vList("ActionNumber") = mvActionNumber.ToString
            DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionTopic, vList)
            RefreshCard()
            Return True
          End If
        Case XMLExamDataSelectionTypes.ExamCentreActionLinks
          Dim vValid As Boolean = True
          If mvCurrentEpl.GetValue("ActionLinkType").Length < 1 Then
            mvCurrentEpl.SetErrorField("ActionLinkType", InformationMessages.ImFieldMandatory, True)
            vValid = False
          End If
          If vValid Then vValid = mvCurrentEpl.AddValuesToList(vList, False, EditPanel.AddNullValueTypes.anvtNone)
          If vValid Then
            vList("ActionNumber") = mvActionNumber.ToString
            DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)

            If mvCurrentEpl.Recipients IsNot Nothing AndAlso mvCurrentEpl.Recipients.Rows.Count > 1 Then
              'This is PostPoint contacts - need to add Action Link to each Contact
              'First Contact has already been done
              For vIndex As Integer = 1 To epl.Recipients.Rows.Count - 1
                vList("ContactNumber") = epl.Recipients.Rows(vIndex).Item("ContactNumber").ToString
                DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
              Next
            End If

            RefreshCard()
            Return True
          End If
        Case XMLExamDataSelectionTypes.ExamCentreUnits
          Dim vChangedList As List(Of ChangedItem) = selExams.GetChangedList
          Dim vAddCount As Integer = 0
          Dim vDeleteCount As Integer = 0
          Dim vAddList As New ParameterList
          Dim vDeleteList As New ParameterList
          Dim vBatchLimit As Integer = 500
          For Each vChangedItem In vChangedList
            If vChangedItem.Checked Then
              'Need to add it
              If vAddCount = 0 Then
                vAddList = New ParameterList(True)
                vAddList.IntegerValue("ExamCentreId") = mvExamCentreId
                vAddList.IntegerValue("ExamUnitId") = vChangedItem.Item.UnitID
                vAddList.IntegerValue("ExamUnitLinkId") = vChangedItem.Item.LinkID
                vAddCount = 1
              Else
                vAddList("ExamUnitId") = vAddList("ExamUnitId") & "," & vChangedItem.Item.UnitID.ToString
                vAddList("ExamUnitLinkId") = vAddList("ExamUnitLinkId") & "," & vChangedItem.Item.LinkID.ToString
                vAddCount = vAddCount + 1
              End If
              If vAddCount = vBatchLimit Then
                ExamsDataHelper.AddItem(mvExamDataType, vAddList)
                vAddCount = 0
              End If
            Else
              'Need to remove it
              If vDeleteCount = 0 Then
                vDeleteList = New ParameterList(True)
                vDeleteList.IntegerValue("ExamCentreUnitId") = vChangedItem.ID2
                vDeleteCount = 1
              Else
                vDeleteList("ExamCentreUnitId") = vDeleteList("ExamCentreUnitId") & "," & vChangedItem.ID2.ToString
                vDeleteCount = vDeleteCount + 1
              End If
              If vDeleteCount = vBatchLimit Then
                ExamsDataHelper.DeleteItem(mvExamDataType, vDeleteList)
                vDeleteCount = 0
              End If
            End If
          Next
          If vAddCount > 0 Then
            ExamsDataHelper.AddItem(mvExamDataType, vAddList)
          End If
          If vDeleteCount > 0 Then
            ExamsDataHelper.DeleteItem(mvExamDataType, vDeleteList)
          End If

          'selExams.Init(ExamSelector.SelectionType.CentreCourses, mvExamCentreId)
          'sel.Init(ExamSelector.SelectionType.CentreCourses, mvExamCentreId)
          mvRefreshSelector = (vChangedList IsNot Nothing AndAlso vChangedList.Count > 0)
          Return True
        Case XMLExamDataSelectionTypes.ExamExemptionUnits
          Dim vUnitID As Integer = selExams.GetUnitID
          If vUnitID > 0 Then
            vList.IntegerValue("ExamExemptionId") = mvExamExemptionId
            vList.IntegerValue("ExamUnitId") = vUnitID
            ExamsDataHelper.AddItem(mvExamDataType, vList)
          End If
          selExams.Init(ExamSelector.SelectionType.ExemptionCourses, mvExamExemptionId)
          Return True
        Case XMLExamDataSelectionTypes.ExamSessionCentres
          Dim vChangedList As List(Of ChangedItem) = selExams.GetChangedList
          For Each vChangedItem In vChangedList
            vList = New ParameterList(True)
            If vChangedItem.Checked Then
              vList.IntegerValue("ExamSessionId") = mvExamSessionId
              vList.IntegerValue("ExamCentreId") = vChangedItem.ID
              ExamsDataHelper.AddItem(mvExamDataType, vList)
              'Need to add it
            Else
              vList.IntegerValue("ExamSessionCentreId") = vChangedItem.ID2
              ExamsDataHelper.DeleteItem(mvExamDataType, vList)
              'Need to remove it
            End If
          Next
          selExams.Init(ExamSelector.SelectionType.SessionCentres, mvExamSessionId)
        Case Else
          If mvExamDataType = XMLExamDataSelectionTypes.ExamSchedule AndAlso mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then
            Dim vChangedList As List(Of ChangedItem) = mvExamScheduleSelector.GetChangedList
            Dim vScheduleValid As Boolean = mvCurrentEpl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)
            vList.Remove("UpdateSchedules")

            If vScheduleValid Then
              For Each vChangedItem In vChangedList
                If vChangedItem.Checked Then
                  ' Insert new schedules
                  vList.IntegerValue("ExamSessionId") = mvCourseSessionId
                  vList.IntegerValue("ExamUnitId") = mvExamUnitId
                  vList.IntegerValue("ExamCentreId") = vChangedItem.ID
                  ExamsDataHelper.AddItem(XMLExamMaintenanceTypes.ExamSchedule, vList)
                Else
                  ' Delete old schedules
                  vList.IntegerValue("ExamScheduleId") = vChangedItem.ID2
                  ExamsDataHelper.DeleteItem(XMLExamMaintenanceTypes.ExamSchedule, vList)
                End If
              Next
              Dim vCreatedSchedules As Integer = vChangedList.Where(Function(item) item.Checked).Count()
              Dim vDeletedSchedules As Integer = vChangedList.Count - vCreatedSchedules
              Dim vCheckedList As List(Of ChangedItem) = mvExamScheduleSelector.GetCheckedList
              Dim vExistingSchedules As Integer = vCheckedList.Where(Function(checkedItem) checkedItem.ID2 > 0).Count()
              'Check if any checked items are existing schedules, i.e. they have a Schedule ID
              If vExistingSchedules > 0 Then
                'Build the messsage with the values that are on the screen (based on possible customisation)
                Dim vMsg As String = QuestionMessages.QmUpdateExistingExamShedules 'BR20058 - added values to message
                If Not String.IsNullOrEmpty(vMsg) AndAlso vMsg.Contains("{3}") Then
                  Dim vValues As String = ""
                  Dim vValueSeparator = ""
                  For Each vKey As Object In vList.Keys
                    Dim vParam As Object = vKey.ToString()
                    If Not String.IsNullOrEmpty(vList(vKey).ToString()) Then
                      Dim vControl As TransparentLabel = mvCurrentEpl.FindLabel(vParam.ToString(), False)
                      If vControl IsNot Nothing Then
                        vValues += String.Format("{0}{1}: {2}", vValueSeparator, vControl.Text, vList(vKey).ToString())
                        vValueSeparator = ", "
                      End If
                    End If
                  Next
                  vMsg = String.Format(vMsg, vCreatedSchedules, vDeletedSchedules, vExistingSchedules, vValues)
                End If

                'Ask user if they want to perform update of existing schedules with date and time, or did they just want to do inserts and deletes (done above)
                If ShowQuestion(vMsg, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then 'BR20058 - changed message
                  For Each vCheckedItem As ChangedItem In vCheckedList
                    'Ensure that the current item has an ExamScheduleId and that it is not one of the ones that was inserted
                    If vCheckedItem.ID2 > 0 Then
                      vList.IntegerValue("ExamScheduleId") = vCheckedItem.ID2
                      vList.IntegerValue("ExamSessionId") = mvCourseSessionId
                      vList.IntegerValue("ExamUnitId") = mvExamUnitId
                      vList.IntegerValue("ExamCentreId") = vCheckedItem.ID
                      ExamsDataHelper.UpdateItem(XMLExamMaintenanceTypes.ExamSchedule, vList)
                    End If
                  Next
                End If
              End If
              Return True
            Else
              Return False
            End If
          End If
          Dim vEditing As Boolean = EditingExistingRecord()
          If vEditing Then
            'If editing an existing record then get the primary key values
            GetPrimaryKeyValues(vList, mvSelectedRow, True)
          Else
            'For new records add in any additional key values
            GetAdditionalKeyValues(vList)
            If mvParentID > 0 AndAlso mvExamDataType = XMLExamDataSelectionTypes.ExamCentres Then
              vList.IntegerValue("ExamCentreParentId") = mvParentID
            End If
          End If

          Dim vValid As Boolean = mvCurrentEpl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)

          If (mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentres OrElse
            mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnit OrElse
            mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreUnitDetails) Then

            Dim vAccreditationFromDate As String = String.Empty
            Dim vAccreditationToDate As String = String.Empty

            If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreUnitDetails Then
              vAccreditationFromDate = mvCurrentEpl.GetValue("CourseAccreditationValidFrom")
              vAccreditationToDate = mvCurrentEpl.GetValue("CourseAccreditationValidTo")
            Else
              vAccreditationFromDate = mvCurrentEpl.GetValue("AccreditationValidFrom")
              vAccreditationToDate = mvCurrentEpl.GetValue("AccreditationValidTo")
            End If

            If (vAccreditationToDate.Length > vAccreditationFromDate.Length) OrElse
              (vAccreditationFromDate.Length > 0 And vAccreditationToDate.Length > 0) AndAlso
              (DateValue(vAccreditationFromDate) > DateValue(vAccreditationToDate)) Then
              ShowInformationMessage(InformationMessages.ImInvalidAccDateRange, "Valid From - Valid To")
              vValid = False
            End If

            If vValid Then
              If vList.ContainsKey("AccreditationStatus") AndAlso vList("AccreditationStatus").Length = 0 AndAlso
              ((vList.ContainsKey("AccreditationValidFrom") = True AndAlso vList("AccreditationValidFrom").Length > 0) _
               OrElse (vList.ContainsKey("AccreditationValidTo") = True AndAlso vList("AccreditationValidTo").Length > 0)) Then
                ShowInformationMessage(InformationMessages.ImExamAccreditationStatusRequiredWhenValidFromValidToSet)
                vValid = False
              ElseIf vList.ContainsKey("AccreditationStatus") AndAlso vList("AccreditationStatus").Length > 0 AndAlso
              (vList.ContainsKey("AccreditationValidFrom") = False OrElse vList("AccreditationValidFrom").Length = 0) Then
                mvCurrentEpl.SetErrorField("AccreditationValidFrom", InformationMessages.ImFieldMandatory)
                vValid = False
              Else
                If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidFrom") Then
                  mvCurrentEpl.SetErrorField("AccreditationValidFrom", "")
                  mvCurrentEpl.PanelInfo.PanelItems("AccreditationValidFrom").Mandatory = False
                End If
              End If
            End If

            If vValid Then
              If vList.ContainsKey("CourseAccreditation") AndAlso vList("CourseAccreditation").Length = 0 AndAlso
              ((vList.ContainsKey("CourseAccreditationValidFrom") = True AndAlso vList("CourseAccreditationValidFrom").Length > 0) _
               OrElse (vList.ContainsKey("CourseAccreditationValidTo") = True AndAlso vList("CourseAccreditationValidTo").Length > 0)) Then
                ShowInformationMessage(InformationMessages.ImExamAccreditationStatusRequiredWhenValidFromValidToSet)
                vValid = False
              ElseIf vList.ContainsKey("CourseAccreditation") AndAlso vList("CourseAccreditation").Length > 0 AndAlso
              (vList.ContainsKey("CourseAccreditationValidFrom") = False OrElse vList("CourseAccreditationValidFrom").Length = 0) Then
                mvCurrentEpl.SetErrorField("CourseAccreditationValidFrom", InformationMessages.ImFieldMandatory)
                vValid = False
              Else
                If mvCurrentEpl.PanelInfo.PanelItems.Exists("CourseAccreditationValidFrom") Then
                  mvCurrentEpl.SetErrorField("CourseAccreditationValidFrom", "")
                  mvCurrentEpl.PanelInfo.PanelItems("CourseAccreditationValidFrom").Mandatory = False
                End If
              End If
            End If
          End If

          If vValid Then
            If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamSchedule Then
              Dim vCentreId As Integer = mvCurrentEpl.FindTextLookupBox("ExamCentreCode").GetDataRowInteger("ExamCentreId")
              vList.IntegerValue("ExamCentreId") = vCentreId
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitGrades Then
              Select Case vList("ExamGradeConditionType")
                Case "G"
                  vList("RequiredValue") = vList("RequiredGrade")
                Case "P"
                  vList("RequiredValue") = vList("RequiredResult")
                Case "S"
                  vList("RequiredValue") = vList("RequiredStatus")
                Case "M"
                  vList("RequiredValue") = vList("RequiredMark")
              End Select
              vList.Remove("RequiredGrade")
              vList.Remove("RequiredResult")
              vList.Remove("RequiredStatus")
              vList.Remove("RequiredMark")
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitPrerequisites Then
              vList("ExamUnitCode") = vList("ExamPrerequisiteUnitCode")
              If vEditing Then
                vList.IntegerValue("OldExamPrerequisiteUnitId") = vList.IntegerValue("ExamPrerequisiteUnitId")
              End If
              vList.IntegerValue("ExamPrerequisiteUnitId") = mvCurrentEpl.FindTextLookupBox("ExamPrerequisiteUnitCode").GetDataRowInteger("ExamUnitId")
              vList.IntegerValue("ExamUnitId") = mvExamUnitId
              'vList("MinimumGrade") = mvCurrentEpl.FindTextLookupBox("ExamGrade").Text
              vList.Remove("ExamPrerequisiteUnitCode")
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnit Then
              vList.IntegerValue("ExamUnitLinkId") = mvExamSelectorItem.LinkID
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitLinkCategories Then
              vList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreCategories Then
              vList.IntegerValue("ExamCentreId") = mvExamCentreId
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories Then
              vList.IntegerValue("ExamCentreUnitId") = mvExamSelectorItem.CentreUnitID
            ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreUnitDetails Then
              vList.IntegerValue("ExamCentreUnitId") = mvExamSelectorItem.CentreUnitID
            End If


            Dim vReturnList As ParameterList
            If vEditing Then
              If ConfirmUpdate() = False Then Exit Function
              vReturnList = ExamsDataHelper.UpdateItem(mvExamMaintenanceType, vList)
              'Update the tree view item
              Dim vDescItem As String = ""
              Select Case mvExamDataType
                Case XMLExamDataSelectionTypes.ExamCentres
                  vDescItem = "ExamCentreDescription"
                Case XMLExamDataSelectionTypes.ExamExemptions
                  vDescItem = "ExamExemptionDescription"
                Case XMLExamDataSelectionTypes.ExamSessions
                  vDescItem = "ExamSessionDescription"
                Case XMLExamDataSelectionTypes.ExamUnits
                  vDescItem = "ExamUnitDescription"
                Case XMLExamDataSelectionTypes.ExamCentreUnitDetails
                  If vList.ContainsKey("LocalName") AndAlso vList("LocalName").Length > 0 Then
                    sel.UpdateSelectedNodeText(vList("LocalName") + "(" + vList("ExamUnitCode") + ")")
                  Else
                    vDescItem = "ExamUnitDescription"
                  End If
              End Select
              If vList.ContainsKey(vDescItem) Then sel.UpdateSelectedNodeText(vList(vDescItem))
            Else
              If ConfirmInsert() = False Then Exit Function
              vReturnList = ExamsDataHelper.AddItem(mvExamMaintenanceType, vList)
              'Remember the primary key to select in the treeview
              Select Case mvExamDataType
                Case XMLExamDataSelectionTypes.ExamCentres
                  mvExamCentreId = vReturnList.IntegerValue("ExamCentreId")
                  Dim vCreateUnitsList As New ParameterList(True)
                  vCreateUnitsList.IntegerValue("ExamCentreId") = mvExamCentreId
                  ExamsDataHelper.CreateExamCentreUnits(vCreateUnitsList)
                Case XMLExamDataSelectionTypes.ExamExemptions
                  mvExamExemptionId = vReturnList.IntegerValue("ExamExemptionId")
                Case XMLExamDataSelectionTypes.ExamSessions
                  Dim vCreateCentresList As New ParameterList(True)
                  If IsCloning Then
                    vCreateCentresList.IntegerValue("CloneId") = mvExamSessionId
                  End If
                  mvExamSessionId = vReturnList.IntegerValue("ExamSessionId")
                  vCreateCentresList.IntegerValue("ExamSessionId") = mvExamSessionId
                  ExamsDataHelper.CreateExamSessionCentres(vCreateCentresList)
                  IsCloning = False
                Case XMLExamDataSelectionTypes.ExamUnits
                  mvExamUnitId = vReturnList.IntegerValue("ExamUnitId")
                  'If mvParentID > 0 Then
                  Dim vLinkList As New ParameterList(True)
                  vLinkList.IntegerValue("ExamUnitId1") = mvParentID
                  vLinkList.IntegerValue("ExamUnitId2") = mvExamUnitId
                  'vLinkList.IntegerValue("ParentUnitLinkId") = mvItemInfo.IntegerValue("ParentLinkID")
                  If mvCurrentEpl.PanelInfo.PanelItems.Exists("LongDescription") Then vLinkList("LongDescription") = mvCurrentEpl.GetValue("LongDescription")
                  If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationStatus") Then vLinkList("AccreditationStatus") = mvCurrentEpl.GetValue("AccreditationStatus")
                  If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidFrom") Then vLinkList("AccreditationValidFrom") = mvCurrentEpl.GetValue("AccreditationValidFrom")
                  If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidTo") Then vLinkList("AccreditationValidTo") = mvCurrentEpl.GetValue("AccreditationValidTo")
                  vLinkList.IntegerValue("ParentUnitLinkId") = mvParentLinkID
                  Dim vLinkReturnList As ParameterList = ExamsDataHelper.AddItem(XMLExamDataSelectionTypes.ExamUnitLinks, vLinkList)
                  mvExamUnitLinkId = vLinkReturnList.IntegerValue("ExamUnitLinkId")
                  'End If
                Case XMLExamDataSelectionTypes.ExamCentreCategories, XMLExamDataSelectionTypes.ExamCentreUnitLinkCategories, XMLExamDataSelectionTypes.ExamUnitLinkCategories
                  RefreshCard()
              End Select
              mvRefreshSelector = True
            End If
            If mvExamDataType = XMLExamDataSelectionTypes.ExamSchedule AndAlso
               vList.ContainsKey("ExamScheduleId") Then
              Dim vParams As New ParameterList(True, True)
              vParams("ExamScheduleId") = vList("ExamScheduleId")
              Dim vForm As New frmApplicationParameters(FunctionParameterTypes.fptExamScheduleWorkstreams, vParams, Nothing)
              If vForm.IsFormValid Then
                AddHandler vForm.OpenWorkstreamGroup, Sub(pWorkstreamGroup As String, pWorkstreams As IList(Of Integer))
                                                        FormHelper.ShowWorkstreamIndex(pWorkstreamGroup, pWorkstreams)
                                                      End Sub
                vForm.Show()
              Else
                vForm.Dispose()
              End If
            End If
            mvCurrentEpl.DataChanged = False     'Data saved now
            Return True
          End If
      End Select
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          Select Case mvExamDataType
            Case XMLExamDataSelectionTypes.ExamExemptionUnits
              ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
            Case Else
              Select Case mvExamMaintenanceType
                Case XMLExamMaintenanceTypes.ExamCentres
                  mvCurrentEpl.SetErrorField("ExamCentreCode", InformationMessages.ImRecordAlreadyExists)
                Case XMLExamMaintenanceTypes.ExamExemptions
                  mvCurrentEpl.SetErrorField("ExamExemptionCode", InformationMessages.ImRecordAlreadyExists)
                Case XMLExamMaintenanceTypes.ExamPersonnel
                  mvCurrentEpl.SetErrorField("ExamPersonnelType", InformationMessages.ImRecordAlreadyExists)
                Case XMLExamMaintenanceTypes.ExamSessions
                  mvCurrentEpl.SetErrorField("ExamSessionCode", InformationMessages.ImRecordAlreadyExists)
                Case XMLExamMaintenanceTypes.ExamUnit
                  mvCurrentEpl.SetErrorField("ExamUnitCode", InformationMessages.ImRecordAlreadyExists)
                Case XMLExamMaintenanceTypes.ExamUnitPersonnel
                  mvCurrentEpl.SetErrorField("ExamPersonnelId", InformationMessages.ImRecordAlreadyExists)
                Case XMLExamMaintenanceTypes.ExamUnitResources
                  mvCurrentEpl.SetErrorField("Product", InformationMessages.ImRecordAlreadyExists)
                Case Else
                  ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
              End Select
          End Select
        Case CareException.ErrorNumbers.enActivityAlreadyExistsInTimePeriod, CareException.ErrorNumbers.enActivityAlreadyExists,
          CareException.ErrorNumbers.enParameterMissing, CareException.ErrorNumbers.enCannotDeleteECUDocsExist
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enExamBooking
          mvCurrentEpl.SetErrorField("GradeUnits", vEx.Message)
          'ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enCannotDelete
          Select Case mvExamDataType
            Case XMLExamDataSelectionTypes.ExamSessionCentres
              selExams.SelectNode(vChangedItem.ID, False, True, XMLExamDataSelectionTypes.ExamSessionCentres)
          End Select
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enExamCentreValidFromBeforeParentValidFrom, CareException.ErrorNumbers.enExamCentreValidToAfterParentValidTo,
             CareException.ErrorNumbers.enExamCentreValidFromAfterChildValidFrom, CareException.ErrorNumbers.enExamCentreValidToBeforeChildValidTo,
             CareException.ErrorNumbers.enExamCentreValidFromBeforeParentValidFromAndNull, CareException.ErrorNumbers.enExamCentreValidToAfterParentValidToAndNull,
             CareException.ErrorNumbers.enExamCentreValidFromAfterChildValidFromAndNull, CareException.ErrorNumbers.enExamCentreValidToBeforeChildValidToAndNull,
             CareException.ErrorNumbers.enCannotDeleteECUDocsExist, CareException.ErrorNumbers.enInvalidActionDuration
          ShowInformationMessage(vEx.Message)
        Case Else
          Throw vEx
      End Select
    End Try
  End Function

  Private Sub GetPrimaryKeyValues(ByVal pList As ParameterList, ByVal pRow As Integer, ByVal pForUpdate As Boolean)
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamCentreAssessmentTypes
        pList("ExamCentreAssessmentTypeId") = dgr.GetValue(pRow, "ExamCentreAssessmentTypeId")
      Case XMLExamMaintenanceTypes.ExamCentres
        pList.IntegerValue("ExamCentreId") = mvExamCentreId
      Case XMLExamMaintenanceTypes.ExamCentreContacts
        pList("ExamCentreContactId") = dgr.GetValue(pRow, "ExamCentreContactId")
      Case XMLExamMaintenanceTypes.ExamCentreActions
        pList("ActionNumber") = dgr.GetValue(pRow, "ActionNumber")
        pList.IntegerValue("ExamCentreId") = mvExamCentreId
      Case XMLExamMaintenanceTypes.ExamExemptions
        pList.IntegerValue("ExamExemptionId") = mvExamExemptionId
      Case XMLExamMaintenanceTypes.ExamPersonnel
        pList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
      Case XMLExamMaintenanceTypes.ExamPersonnelAssessmentTypes
        pList("ExamPersonnelAssessTypeId") = dgr.GetValue(pRow, "ExamPersonnelAssessTypeId")
      Case XMLExamMaintenanceTypes.ExamPersonnelExpenses
        pList("ExamPersonnelExpenseId") = dgr.GetValue(pRow, "ExamPersonnelExpenseId")
      Case XMLExamMaintenanceTypes.ExamSchedule
        pList("ExamScheduleId") = dgr.GetValue(pRow, "ExamScheduleId")
      Case XMLExamMaintenanceTypes.ExamSessions
        pList.IntegerValue("ExamSessionId") = mvExamSessionId
      Case XMLExamMaintenanceTypes.ExamUnit
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
      Case XMLExamMaintenanceTypes.ExamUnitAssessmentTypes
        pList("ExamUnitAssessmentTypeId") = dgr.GetValue(pRow, "ExamUnitAssessmentTypeId")
      Case XMLExamMaintenanceTypes.ExamUnitGrades
        pList("ExamUnitGradeId") = dgr.GetValue(pRow, "ExamUnitGradeId")
      Case XMLExamMaintenanceTypes.ExamUnitPersonnel
        pList("ExamUnitPersonnelId") = dgr.GetValue(pRow, "ExamUnitPersonnelId")
      Case XMLExamMaintenanceTypes.ExamUnitPrerequisites
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
        pList("ExamPrerequisiteUnitId") = dgr.GetValue(pRow, "ExamPrerequisiteUnitId")
      Case XMLExamMaintenanceTypes.ExamUnitRequirements
        pList("ExamUnitEligibilityCheckId") = dgr.GetValue(pRow, "ExamUnitEligibilityCheckId")
      Case XMLExamMaintenanceTypes.ExamUnitResources
        pList("ExamUnitProductId") = dgr.GetValue(pRow, "ExamUnitProductId")
      Case XMLExamMaintenanceTypes.ExamUnitMarkerAllocation
        If dgr.GetValue(pRow, "ExamPersonnelId").Length > 0 Then pList("ExamPersonnelId") = dgr.GetValue(pRow, "ExamPersonnelId")
        pList("ExamUnitId") = dgr.GetValue(pRow, "ExamUnitId")
        If dgr.GetValue(pRow, "MarkerNumber").Length > 0 Then pList("MarkerNumber") = dgr.GetValue(pRow, "MarkerNumber")
      Case XMLExamMaintenanceTypes.ExamUnitLinkCategories
        If dgr.GetValue(pRow, "CategoryId").Length > 0 Then pList("CategoryId") = dgr.GetValue(pRow, "CategoryId")
        If dgr.GetValue(pRow, "ExamUnitLinkId").Length > 0 Then pList("ExamUnitLinkId") = dgr.GetValue(pRow, "ExamUnitLinkId")
      Case XMLExamMaintenanceTypes.ExamCentreCategories
        If dgr.GetValue(pRow, "ExamCentreId").Length > 0 Then pList("ExamCentreId") = dgr.GetValue(pRow, "ExamCentreId")
        If dgr.GetValue(pRow, "CategoryId").Length > 0 Then pList("CategoryId") = dgr.GetValue(pRow, "CategoryId")
      Case XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories
        If dgr.GetValue(pRow, "ExamCentreUnitId").Length > 0 Then pList("ExamCentreUnitId") = dgr.GetValue(pRow, "ExamCentreUnitId")
        If dgr.GetValue(pRow, "CategoryId").Length > 0 Then pList("CategoryId") = dgr.GetValue(pRow, "CategoryId")
        If dgr.GetValue(pRow, "ExamUnitLinkId").Length > 0 Then pList("ExamUnitLinkId") = dgr.GetValue(pRow, "ExamUnitLinkId")
      Case XMLExamMaintenanceTypes.ExamUnitLinkDocuments, XMLExamMaintenanceTypes.ExamCentreDocuments, XMLExamMaintenanceTypes.ExamCentreUnitLinkDocuments
        pList("DocumentNumber") = dgr.GetValue(pRow, "DocumentNumber")
      Case XMLExamMaintenanceTypes.ExamUnitCertRunTypes
        pList.IntegerValue("ExamUnitLinkId") = sel.GetLinkID
        pList("ExamCertRunType") = dgr.GetValue(pRow, "ExamCertRunType")
      Case XMLExamMaintenanceTypes.ExamCentreActionLinks
        pList("ContactNumber") = dgr.GetValue(pRow, "ContactNumber")

      Case Else
        Select Case mvExamDataType
          Case XMLExamDataSelectionTypes.ExamExemptionUnits
            pList("ExamExemptionUnitId") = dgr.GetValue(pRow, "ExamExemptionUnitId")
          Case XMLExamDataSelectionTypes.ExamPersonnelMarkerInfo
            pList("ExamPersonnelId") = dgr.GetValue(pRow, "ExamPersonnelId")
            pList("ExamUnitId") = dgr.GetValue(pRow, "ExamUnitId")
            pList("MarkerNumber") = dgr.GetValue(pRow, "MarkerNumber")
            pList("ExamCentreId") = dgr.GetValue(pRow, "ExamCentreId")
        End Select
    End Select
    Select Case mvExamDataType
      Case XMLExamDataSelectionTypes.ExamExemptionUnits
        pList("ExamExemptionUnitId") = dgr.GetValue(pRow, "ExamExemptionUnitId")
    End Select
  End Sub

  Private Sub GetAdditionalKeyValues(ByVal pList As ParameterList)
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamCentreAssessmentTypes
        pList.IntegerValue("ExamCentreId") = mvExamCentreId
      Case XMLExamMaintenanceTypes.ExamCentres
        'pList.IntegerValue("ExamCentreId") = mvExamCentreId
      Case XMLExamMaintenanceTypes.ExamCentreContacts
        pList.IntegerValue("ExamCentreId") = mvExamCentreId
      Case XMLExamMaintenanceTypes.ExamPersonnel
        'pList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
      Case XMLExamMaintenanceTypes.ExamPersonnelAssessmentTypes
        pList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
      Case XMLExamMaintenanceTypes.ExamPersonnelExpenses
        pList.IntegerValue("ExamPersonnelId") = mvExamPersonnelId
      Case XMLExamMaintenanceTypes.ExamSchedule
        pList.IntegerValue("ExamSessionId") = mvCourseSessionId
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
      Case XMLExamMaintenanceTypes.ExamSessions
        'pList.IntegerValue("ExamSessionId") = mvExamSessionId
      Case XMLExamMaintenanceTypes.ExamUnit
        If mvCourseSessionId > 0 Then pList.IntegerValue("ExamSessionId") = mvCourseSessionId
      Case XMLExamMaintenanceTypes.ExamUnitAssessmentTypes
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
      Case XMLExamMaintenanceTypes.ExamUnitGrades
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
      Case XMLExamMaintenanceTypes.ExamUnitPersonnel
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
      Case XMLExamMaintenanceTypes.ExamUnitRequirements
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
      Case XMLExamMaintenanceTypes.ExamUnitResources
        pList.IntegerValue("ExamUnitId") = mvExamUnitId
    End Select
  End Sub

  Private Function EditingExistingRecord() As Boolean
    If mvSelectedRow >= 0 Then Return True
  End Function

  Private Sub RePopulateGrid()
    RefreshCard()
  End Sub

  Private Sub ProcessNew()
    dgr.SelectRow(-1)
    mvSelectedRow = -1
    mvCurrentEpl.Clear()
    SetDefaults(False)
    SetCommandsForNew()
    mvParentID = 0
    mvParentLinkID = 0
    IsCloning = False
  End Sub

  Private Sub ProcessClone()
    dgr.SelectRow(-1)
    mvSelectedRow = -1
    mvCurrentEpl.SetValue("ValidFrom", AppValues.TodaysDate)
    mvCurrentEpl.SetValue("ExamSessionCode", String.Empty)
    mvCurrentEpl.EnableControlList("ExamSessionCode", True)
    SetCommandsForNew()
    mvParentID = 0
    mvParentLinkID = 0
    IsCloning = True
  End Sub

  Protected Overridable Sub SetDefaults(Optional ByVal pInitialSetup As Boolean = True)
    If dgrDetails.Visible Then dgrDetails.Clear()
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamExemptions
        mvCurrentEpl.EnableControlList("ExamExemptionCode", True)
      Case XMLExamMaintenanceTypes.ExamCentres
        mvCurrentEpl.EnableControl("ExamCentreCode", True)
        If mvCurrentEpl.PanelInfo IsNot Nothing AndAlso
          mvCurrentEpl.PanelInfo.PanelItems.Exists("ContactNumber") Then
          mvCurrentEpl.FindTextLookupBox("ContactNumber").ClearDataSource()
        End If
        If dspTabGrid.Visible Then
          dspTabGrid.DisplayGrid(0).Clear()
          dspTabGrid.DisplayGrid(1).Clear()
        End If
      Case XMLExamMaintenanceTypes.ExamCentreActions
        mvCurrentEpl.SetValue("DocumentClass", AppValues.DefaultDocumentClass)
        mvCurrentEpl.SetValue("ActionPriority", AppValues.DefaultActionPriority)
        mvCurrentEpl.SetValue("DurationDays", AppValues.DefaultActionDuration.Days.ToString("##", CultureInfo.InvariantCulture))
        mvCurrentEpl.SetValue("DurationHours", AppValues.DefaultActionDuration.Hours.ToString("##", CultureInfo.InvariantCulture))
        mvCurrentEpl.SetValue("DurationMinutes", AppValues.DefaultActionDuration.Minutes.ToString("##", CultureInfo.InvariantCulture))
        SetActionChangeReason(mvCurrentEpl, False, True)
      Case XMLExamMaintenanceTypes.ExamCentreContacts
        mvCurrentEpl.SetValue("ContactNumber", sel.GetParentID2.ToString)
      Case XMLExamMaintenanceTypes.ExamPersonnel
        mvCurrentEpl.SetValue("ValidFrom", AppValues.TodaysDate)
      Case XMLExamMaintenanceTypes.ExamSchedule
        mvCurrentEpl.SetValue("StartTime", "09:00")
      Case XMLExamMaintenanceTypes.ExamSessions
        mvCurrentEpl.EnableControlList("ExamSessionCode", True)
        mvCurrentEpl.SetValue("ValidFrom", AppValues.TodaysDate)
      Case XMLExamMaintenanceTypes.ExamUnit
        mvCurrentEpl.EnableControlList("ExamUnitCode,ExamUnitType", True)
        mvCurrentEpl.SetValue("ExamMarkType", "M")
        mvCurrentEpl.SetValue("MarkFactor", "1")
        mvCurrentEpl.SetValue("SessionBased", "Y")
        mvCurrentEpl.SetValue("AllowExemptions", "Y")
        mvCurrentEpl.SetValue("ExamMarkerStatus", "N")
        mvCurrentEpl.SetValue("TimeLimitType", "N")
        If FindControl(mvCurrentEpl, "IsGradingEndpoint", False) IsNot Nothing Then mvCurrentEpl.SetValue("IsGradingEndpoint", "N")
        If mvCourseSessionId > 0 Then
          Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamUnitTypes)
          For Each vRow As DataRow In vTable.Rows
            If vRow("ExamQuestion").ToString = "Y" Then
              mvCurrentEpl.SetValue("ExamUnitType", vRow("ExamUnitType").ToString, False, True) 'default to first question type but let the user choose.  Validation will stop them from saving non-question types
              Exit For
            End If
          Next
        End If
        If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidFrom") Then
          mvCurrentEpl.SetValue("AccreditationValidFrom", AppValues.TodaysDate)
          mvCurrentEpl.FindDateTimePicker("AccreditationValidFrom").Checked = False
        End If
        If mvCurrentEpl.PanelInfo.PanelItems.Exists("AccreditationValidTo") Then
          mvCurrentEpl.SetValue("AccreditationValidTo", AppValues.TodaysDateAddYears(100))
          mvCurrentEpl.FindDateTimePicker("AccreditationValidTo").Checked = False
        End If
        If mvCurrentEpl.PanelInfo.PanelItems.Exists("CourseAccreditationValidFrom") Then
          mvCurrentEpl.SetValue("CourseAccreditationValidFrom", AppValues.TodaysDate)
          mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidFrom").Checked = False
        End If
        If mvCurrentEpl.PanelInfo.PanelItems.Exists("CourseAccreditationValidTo") Then
          mvCurrentEpl.SetValue("CourseAccreditationValidTo", AppValues.TodaysDateAddYears(100))
          mvCurrentEpl.FindDateTimePicker("CourseAccreditationValidTo").Checked = False
        End If
      Case XMLExamMaintenanceTypes.ExamUnitGrades
        SetGradingRuleRequiredUI("", False)
      Case XMLExamMaintenanceTypes.ExamUnitPersonnel
        mvCurrentEpl.SetValue("ValidFrom", AppValues.TodaysDate)
        mvCurrentEpl.FindTextLookupBox("ExamPersonnelId").ActiveOnly = True
      Case XMLExamMaintenanceTypes.ExamCentreCategories, XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories, XMLExamMaintenanceTypes.ExamUnitLinkCategories
        mvCurrentEpl.SetValue("ValidFrom", AppValues.TodaysDate)
        mvCurrentEpl.SetValue("ValidTo", AppValues.TodaysDateAddYears(100))
        mvCurrentEpl.SetValue("Source", AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_activity_source))
        mvCurrentEpl.EnableControls(True)
        cmdSave.Enabled = True
      Case XMLExamMaintenanceTypes.ExamUnitLinkDocuments
        If Not pInitialSetup Then
          Dim vContactInfo As ContactInfo = DataHelper.UserContactInfo
          vContactInfo.AddressNumber = vContactInfo.AddressNumber
          Dim vParams As New ParameterList()
          vParams.Add("ExamUnitLinkId", mvExamSelectorItem.LinkID)
          FormHelper.NewDocument(Me, vContactInfo, vParams)
          dpl.Visible = True
          'gridToolBar.Visible = True
        Else
          ShowDocumentDetails(0)
        End If
      Case XMLExamMaintenanceTypes.ExamUnitCertRunTypes 'BR20437
        If Not pInitialSetup AndAlso Not EditingExistingRecord() Then mvCurrentEpl.EnableControl("ExamCertRunType", True)
      Case XMLExamMaintenanceTypes.ExamCentreActionLinks
        epl.SetEntityLinkDefaults("ActionLinkType", "R")
    End Select
    mvCurrentEpl.SetUserDefaults()
    mvCurrentEpl.DataChanged = False
  End Sub


  Protected Overridable Sub SetCommandsForNew()
    cmdDelete.Enabled = False
  End Sub

#End Region

#Region "Button Handling"

  Private Sub cmdSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectAll.Click
    Try
      If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then
        mvExamScheduleSelector.SelectAllNodes()
      ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitMarkerAllocation Then
        Dim vCol As Integer = dgrDetails.GetColumn(mvSelectColumnName)
        For vRow As Integer = 0 To dgrDetails.RowCount - 1
          dgrDetails.SetValue(vRow, vCol, "True")
        Next
      Else
        selExams.SelectAllNodes()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdUnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnSelectAll.Click
    Try
      If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then
        mvExamScheduleSelector.UnSelectAllNodes()
      ElseIf mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamUnitMarkerAllocation Then
        Dim vCol As Integer = dgrDetails.GetColumn(mvSelectColumnName)
        For vRow As Integer = 0 To dgrDetails.RowCount - 1
          dgrDetails.SetValue(vRow, vCol, "")
        Next
      Else
        selExams.UnSelectAllNodes()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overridable Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click, cmdNewChild.Click, mvActionMenu.NewAction
    Try
      If mvCurrentEpl.DataChanged Then
        If ConfirmSave() Then
          If ProcessSave(sender) = False Then Exit Sub
        End If
      End If
      'Clear selection on display grid and set defaults for new record
      If mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions Then
        mvActionNumber = 0
      End If
      ProcessNew()
      If sender Is cmdNewChild Then
        Select Case mvExamDataType
          Case XMLExamDataSelectionTypes.ExamCentres
            mvParentID = mvExamCentreId
            mvCurrentEpl.SetValue("OrganisationNumber", mvCentreOrganisation.ToString)
          Case XMLExamDataSelectionTypes.ExamUnits
            mvParentID = mvExamUnitId
            mvParentLinkID = mvExamUnitLinkId
        End Select
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    Try
      If ProcessSave(sender) Then
        If dgr.Visible Or mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then RePopulateGrid()
        If mvRefreshSelector Then
          Select Case mvExamDataType
            Case XMLExamDataSelectionTypes.ExamCentres
              sel.Init(mvSelectorType)
              sel.SelectNode(mvExamCentreId, False, False, XMLExamDataSelectionTypes.ExamCentres)
            Case XMLExamDataSelectionTypes.ExamCentreUnits
              Dim vLastNode As TreeNode = selExams.LastCheckedUncheckedNode
              If vLastNode IsNot Nothing AndAlso vLastNode.Tag IsNot Nothing AndAlso vLastNode.Tag.GetType = GetType(ExamSelectorItem) Then
                Dim vLinkID As Integer = DirectCast(vLastNode.Tag, ExamSelectorItem).LinkID
                sel.Init(SelectionType.Centres)
                sel.SelectNode(mvExamCentreId, True, False, XMLExamDataSelectionTypes.ExamCentres, XMLExamDataSelectionTypes.ExamCentreUnits)
                selExams.Init(ExamSelector.SelectionType.CentreCourses, mvExamCentreId)
                If vLinkID > 0 Then selExams.SelectNode(vLinkID, False, False, XMLExamDataSelectionTypes.ExamCentreUnits)
              End If
            Case XMLExamDataSelectionTypes.ExamExemptions
              sel.Init(mvSelectorType)
              sel.SelectNode(mvExamExemptionId, False, False, XMLExamDataSelectionTypes.ExamExemptions)
            Case XMLExamDataSelectionTypes.ExamPersonnel
              sel.Init(mvSelectorType)
              sel.SelectNode(mvExamPersonnelId, False, False, XMLExamDataSelectionTypes.ExamPersonnel)
            Case XMLExamDataSelectionTypes.ExamSessions
              sel.Init(mvSelectorType)
              sel.SelectNode(mvExamSessionId, False, False, XMLExamDataSelectionTypes.ExamSessions)
            Case XMLExamDataSelectionTypes.ExamUnits
              sel.Init(mvSelectorType, mvCourseSessionId)
              sel.SelectNode(mvExamUnitLinkId, False, False, XMLExamDataSelectionTypes.ExamUnits)
          End Select
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overridable Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    If mvSelectedRow <> -1 Then
      Try
        Dim vList As New ParameterList(True)
        GetPrimaryKeyValues(vList, mvSelectedRow, False)
        If Settings.ConfirmDelete AndAlso ShowQuestion(QuestionMessages.QmConfirmDelete, MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.Cancel Then Exit Sub
        Select Case mvExamDataType
          Case XMLExamDataSelectionTypes.ExamCentreActionLinks
            vList("ContactNumber") = dgr.GetValue(mvSelectedRow, "ContactNumber")
            Select Case dgr.GetValue(mvSelectedRow, "EntityType").ToUpper
              Case "D"
                'Document
                vList("DocumentNumber") = vList("ContactNumber")
                vList.Remove("ContactNumber")
              Case "N"
                'Exam Centre
                vList("ExamCentreId") = vList("ContactNumber")
                vList.Remove("ContactNumber")
              Case Else
                'Contact / Organisation
            End Select
            vList.IntegerValue("ActionNumber") = mvActionNumber
            vList("ActionLinkType") = CStr(dgr.DataSourceDataRow(dgr.CurrentDataRow)("LinkType"))
          Case XMLExamDataSelectionTypes.ExamCentreActionAnalysis
            vList("Topic") = CStr(dgr.DataSourceDataRow(dgr.CurrentDataRow)("TopicCode"))
            vList("SubTopic") = CStr(dgr.DataSourceDataRow(dgr.CurrentDataRow)("SubTopicCode"))
            vList.IntegerValue("ActionNumber") = mvActionNumber
        End Select
        If mvExamDataType = XMLExamDataSelectionTypes.ExamExemptionUnits Then
          ExamsDataHelper.DeleteItem(mvExamDataType, vList)
        Else
          ExamsDataHelper.DeleteItem(mvExamMaintenanceType, vList)
        End If
        Select Case mvExamDataType
          Case XMLExamDataSelectionTypes.ExamUnits, XMLExamDataSelectionTypes.ExamCentres, XMLExamDataSelectionTypes.ExamPersonnel, XMLExamDataSelectionTypes.ExamSessions, XMLExamDataSelectionTypes.ExamExemptions
            sel.RemoveSelectedNode()
        End Select
      Catch vEx As CareException
        If vEx.ErrorNumber = CareException.ErrorNumbers.enCannotDelete OrElse
          vEx.ErrorNumber = CareException.ErrorNumbers.enNoAccessRights Then
          ShowInformationMessage(vEx.Message)
        Else
          DataHelper.HandleException(vEx)
        End If
      Catch vException As Exception
        DataHelper.HandleException(vException)
      Finally
        If dgr.Visible Then RePopulateGrid()
      End Try
    End If
  End Sub

  Private Sub cmdAllocate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAllocate.Click
    Try
      Select Case mvExamMaintenanceType
        Case XMLExamMaintenanceTypes.ExamUnitMarkerAllocation
          ' Get list of IDs
          Dim vIdList As New List(Of String)
          Dim vCol As Integer = dgrDetails.GetColumn(mvSelectColumnName)
          For vRow As Integer = 0 To dgrDetails.RowCount - 1
            If dgrDetails.GetValue(vRow, vCol) = "True" Then vIdList.Add(dgrDetails.GetValue(vRow, "ExamBookingUnitId"))
          Next
          If vIdList.Count > 0 Then
            Dim vIdCsvList As String = String.Join(",", vIdList.ToArray())

            If DisplayAllocatePapersToMarker(mvExamUnitId, vIdCsvList, CInt(dgr.GetValue(mvSelectedRow, "MarkerNumber"))) Then
              ' Refresh grids
              Dim vRowCount As Integer = dgr.RowCount
              Dim vRow As Integer = mvSelectedRow
              RePopulateGrid()
              If vRowCount = dgr.RowCount Then dgr.SelectRow(vRow)
            End If
          End If
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function DisplayMarkerSelection(ByVal pList As ParameterList) As Integer
    ' Returns ExamPersonnelId of selected marker. Returns 0 if none selected.
    Dim vExamPersonnelId As Integer = 0

    ' Display list of available markers for selection
    Dim vDataSet As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamMarkerList, pList)
    Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)


    If Not vDataTable Is Nothing Then
      If vDataTable.Rows.Count > 0 Then
        ' Stop the ContactName appearing as a hyper link in the Dialog based grid by renaming it
        If vDataTable.Columns.Contains("ContactName") Then
          vDataTable.Columns("ContactName").ColumnName = "MarkerName"
          DataHelper.GetColumnTableFromDataSet(vDataSet).Select("Name = 'ContactName'")(0).Item("Name") = "MarkerName"
        End If

        Dim vFrmSelect As New frmSelectListItem(vDataSet, frmSelectListItem.ListItemTypes.litExamMarkers)
        If Not vFrmSelect.ShowDialog() = DialogResult.Cancel Then
          vExamPersonnelId = CInt(vDataTable.Rows(vFrmSelect.SelectedRow)("ExamPersonnelId").ToString())
        End If
      Else
        ShowInformationMessage(InformationMessages.ImExamMarkersNotAvailable)
      End If
    Else
      ShowInformationMessage(InformationMessages.ImExamMarkersNotAvailable)
    End If
    Return vExamPersonnelId
  End Function

  Private Function DisplayAllocatePapersToMarker(ByVal pExamUnitId As Integer, ByVal pExamBookingUnitIdList As String, ByVal pMarkerNumber As Integer) As Boolean
    Dim vResult As Boolean = False

    'Dim vMarkerNumber As Integer = DisplayMarkerNumberSelection()
    'If (pMarkerNumber > 0 ) Then

    Dim vList As New ParameterList(True)
    vList.AddSystemColumns()
    vList.Add("ExamBookingUnitId", pExamBookingUnitIdList) ' used to exclude markers already marking the paper
    vList.Add("ExamUnitId", pExamUnitId)

    Dim vExamPersonnelId As Integer = DisplayMarkerSelection(vList)
    If vExamPersonnelId > 0 Then
      vList.Add("ExamPersonnelId", vExamPersonnelId)
      vList.Add("MarkerNumber", pMarkerNumber)

      vResult = AllocatePapersToMarker(vList)
    End If
    'End If
    Return vResult
  End Function

  Private Function AllocatePapersToMarker(ByVal pList As ParameterList) As Boolean
    Dim vResult As Boolean = False

    Dim vResultList As ParameterList = ExamsDataHelper.AddItem(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, pList)
    If vResultList.Contains("Result") Then
      If vResultList("Result") = "OverAllocated" Then
        ' Prompt user for ovverride of Over Allocation check
        Dim vPromptResult As DialogResult = ShowQuestion(QuestionMessages.QmExamMarkerOverAllocated, MessageBoxButtons.OKCancel)
        If vPromptResult = System.Windows.Forms.DialogResult.OK Then
          pList.Add("OverrideMaxPaperCheck", "Y")
          vResultList.Clear()
          vResultList = ExamsDataHelper.AddItem(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, pList)
        End If
      End If
      If vResultList.Contains("Result") AndAlso vResultList("Result") = "OK" Then vResult = True
    End If
    Return vResult
  End Function

  Private Sub cmdUnallocate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdUnallocate.Click
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamUnitMarkerAllocation
        If dgr.DataRowCount > 0 AndAlso dgrDetails.DataSourceDataRow(0) IsNot Nothing Then
          ' Get list of IDs
          Dim vIdList As New List(Of String)
          Dim vCol As Integer = dgrDetails.GetColumn(mvSelectColumnName)
          For vRow As Integer = 0 To dgrDetails.RowCount - 1
            If dgrDetails.GetValue(vRow, vCol) = "True" Then vIdList.Add(dgrDetails.GetValue(vRow, "ExamMarkingBatchDetailId"))
          Next
          If vIdList.Count > 0 Then
            Dim vIdCsvList As String = String.Join(",", vIdList.ToArray())
            Dim vList As New ParameterList(True)
            vList.Add("ExamMarkingBatchDetailId", vIdCsvList)
            ExamsDataHelper.DeleteItem(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, vList)
            Dim vRowCount As Integer = dgr.RowCount
            Dim vRow As Integer = mvSelectedRow
            RePopulateGrid()
            If vRowCount = dgr.RowCount Then dgr.SelectRow(vRow)
          End If
        End If
    End Select
  End Sub

  Private Sub cmdLink_Click(sender As Object, e As EventArgs) Handles cmdLink.Click
    Try
      If mvCurrentEpl.DataChanged Then
        If ConfirmSave() Then
          ProcessSave(sender)
        End If
      End If
      mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionLinks
      RefreshCard()
      SetNodeDataEditable(sel.SelectedNode)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdAnalysis_Click(sender As Object, e As EventArgs) Handles cmdAnalysis.Click
    Try
      If mvCurrentEpl.DataChanged Then
        If ConfirmSave() Then
          ProcessSave(sender)
        End If
      End If
      mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionAnalysis
      RefreshCard()
      SetNodeDataEditable(sel.SelectedNode)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
    mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions
    RefreshCard()
    SetNodeDataEditable(sel.SelectedNode)
  End Sub

#End Region

#Region "Selection Handling"

  Private Sub imgExemptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles imgExemptions.Click
    ProcessSelection(ExamSelector.SelectionType.Exemptions)
  End Sub
  Private Sub imgPersonnel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles imgPersonnel.Click
    Dim vDataChanged As Boolean = mvCurrentEpl.DataChanged OrElse ((mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule) AndAlso (mvExamScheduleSelector IsNot Nothing) AndAlso mvExamScheduleSelector.GetChangedList.Count > 0)
    Dim vCancel As Boolean = False
    If vDataChanged AndAlso ConfirmSave() Then
      vCancel = Not ProcessSave(sender)
    End If
    If Not vCancel Then
      mvCurrentEpl.DataChanged = False
      ProcessSelection(ExamSelector.SelectionType.Personnel)
    End If
  End Sub
  Private Sub imgCourse_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles imgCourse.Click
    Dim vDataChanged As Boolean = mvCurrentEpl.DataChanged OrElse ((mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule) AndAlso (mvExamScheduleSelector IsNot Nothing) AndAlso mvExamScheduleSelector.GetChangedList.Count > 0)
    Dim vCancel As Boolean = False
    If vDataChanged AndAlso ConfirmSave() Then
      vCancel = Not ProcessSave(sender)
    End If
    If Not vCancel Then
      mvCurrentEpl.DataChanged = False
      ProcessSelection(ExamSelector.SelectionType.Courses)
    End If
  End Sub
  Private Sub imgCentres_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles imgCentres.Click
    Dim vDataChanged As Boolean = mvCurrentEpl.DataChanged OrElse ((mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule) AndAlso (mvExamScheduleSelector IsNot Nothing) AndAlso mvExamScheduleSelector.GetChangedList.Count > 0)
    Dim vCancel As Boolean = False
    If vDataChanged AndAlso ConfirmSave() Then
      vCancel = Not ProcessSave(sender)
    End If
    If Not vCancel Then
      mvCurrentEpl.DataChanged = False
      ProcessSelection(ExamSelector.SelectionType.Centres)
    End If
  End Sub
  Private Sub imgSessions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles imgSessions.Click
    Dim vDataChanged As Boolean = mvCurrentEpl.DataChanged OrElse ((mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule) AndAlso (mvExamScheduleSelector IsNot Nothing) AndAlso mvExamScheduleSelector.GetChangedList.Count > 0)
    Dim vCancel As Boolean = False
    If vDataChanged AndAlso ConfirmSave() Then
      vCancel = Not ProcessSave(sender)
    End If
    If Not vCancel Then
      mvCurrentEpl.DataChanged = False
      ProcessSelection(ExamSelector.SelectionType.Sessions)
    End If
  End Sub

  Private Sub ProcessSelection(ByVal pSelectionType As ExamSelector.SelectionType)
    Dim vBusyCursor As New BusyCursor
    Try
      If pSelectionType <> mvSelectorType Then
        imgExemptions.DrawBorder = pSelectionType = ExamSelector.SelectionType.Exemptions
        imgPersonnel.DrawBorder = pSelectionType = ExamSelector.SelectionType.Personnel
        imgCourse.DrawBorder = pSelectionType = ExamSelector.SelectionType.Courses
        imgCentres.DrawBorder = pSelectionType = ExamSelector.SelectionType.Centres
        imgSessions.DrawBorder = pSelectionType = ExamSelector.SelectionType.Sessions

        mvParentID = 0
        mvParentLinkID = 0
        mvSelectorType = pSelectionType
        If pSelectionType = ExamSelector.SelectionType.Courses Then
          PopulateSessions()
          cboSessions.Visible = True
          sel.Init(pSelectionType, mvCourseSessionId)
        Else
          cboSessions.Visible = False
          sel.Init(pSelectionType)
        End If
        sel.SelectNode(0, False)
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Sub sel_BeforeSelect(ByVal sender As Object, ByRef pCancel As Boolean) Handles sel.BeforeSelect
    Dim vDataChanged As Boolean = mvCurrentEpl.DataChanged OrElse ((mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule) AndAlso (mvExamScheduleSelector IsNot Nothing) AndAlso mvExamScheduleSelector.GetChangedList.Count > 0)
    If vDataChanged Then
      If ConfirmSave() Then
        pCancel = Not ProcessSave(sender)
      Else
        'We have some data changed and we are going to cancel it
        mvCurrentEpl.DataChanged = False
      End If
    End If
  End Sub

  Private Sub sel_ItemSelected(ByVal sender As Object, ByVal pType As ExamsAccess.XMLExamDataSelectionTypes, ByVal pItem As ExamSelectorItem) Handles sel.ItemSelected
    Dim vBusyCursor As New BusyCursor()
    Dim vItemID As Integer = 0 'this is the default value if a null ExamSelectorItem is passed
    If pItem IsNot Nothing Then
      vItemID = pItem.ID
      mvExamSelectorItem = pItem
    Else
      pItem = New ExamSelectorItem(XMLExamDataSelectionTypes.ExamUnits, 0, "")
    End If

    Try
      If pType <> XMLExamDataSelectionTypes.Blank Then
        mvExamDataType = pType

        mvSelectedRow = -1
        Select Case mvSelectorType
          Case ExamSelector.SelectionType.Centres
            mvExamCentreId = vItemID
          Case ExamSelector.SelectionType.Courses
            mvExamUnitId = pItem.UnitID
            mvExamUnitLinkId = pItem.LinkID
          Case ExamSelector.SelectionType.Exemptions
            mvExamExemptionId = vItemID
          Case ExamSelector.SelectionType.Personnel
            mvExamPersonnelId = vItemID
          Case ExamSelector.SelectionType.Sessions
            mvExamSessionId = vItemID
        End Select
        RefreshCard()

        SelectRow(mvSelectedRow)

        'Handle setting panels to read-only
        Dim vCanEdit As Boolean = True
        Dim vNode As TreeNode = Nothing
        Select Case mvExamDataType
          'Buttons
          Case XMLExamDataSelectionTypes.ExamCentres
            vCanEdit = Not (imgCentres.Tag.ToString.ToUpper.EndsWith("RO"))
          Case XMLExamDataSelectionTypes.ExamExemptions
            vCanEdit = Not (imgExemptions.Tag.ToString.ToUpper.EndsWith("RO"))
          Case XMLExamDataSelectionTypes.ExamUnits  'Courses
            vCanEdit = Not (imgCourse.Tag.ToString.ToUpper.EndsWith("RO"))
          Case XMLExamDataSelectionTypes.ExamPersonnel
            vCanEdit = Not (imgPersonnel.Tag.ToString.ToUpper.EndsWith("RO"))
          Case XMLExamDataSelectionTypes.ExamSessions
            vCanEdit = Not (imgSessions.Tag.ToString.ToUpper.EndsWith("RO"))
          Case Else
            'Nodes
            vNode = CType(sender, ExamSelector).SelectedNode
        End Select
        If vNode IsNot Nothing Then
          If TypeOf (vNode) Is VistaNode Then vCanEdit = Not (CType(vNode, VistaNode).IsReadOnly)
        End If

        SetNodeDataEditable(vCanEdit)

        If mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActions Or
          mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionLinks Or
          mvExamDataType = XMLExamDataSelectionTypes.ExamCentreActionAnalysis Then
          cmdDelete.Enabled = (dgr.DataRowCount > 0)
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Sub cboSessions_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSessions.SelectedIndexChanged
    If cboSessions.SelectedIndex >= 0 Then
      If IntegerValue(cboSessions.SelectedValue) <> mvCourseSessionId Then
        mvCourseSessionId = IntegerValue(cboSessions.SelectedValue)
        sel.Init(mvSelectorType, mvCourseSessionId)
        sel.SelectNode(0, False)
      End If
    End If
  End Sub

  Private Sub SetNodeDataEditable(ByVal pNode As TreeNode)
    If pNode IsNot Nothing Then
      Dim vCanEdit As Boolean = True
      If TypeOf (pNode) Is VistaNode Then vCanEdit = Not (CType(pNode, VistaNode).IsReadOnly)
      SetNodeDataEditable(vCanEdit)
    End If
  End Sub
  Private Sub SetNodeDataEditable(ByVal pCanEdit As Boolean)
    If pCanEdit = False Then mvCurrentEpl.EnableControls(pCanEdit)
    cmdSave.Enabled = pCanEdit
    cmdNew.Enabled = pCanEdit
    cmdNewChild.Enabled = pCanEdit
    cmdDelete.Enabled = pCanEdit
    cmdAllocate.Enabled = pCanEdit
    cmdSelectAll.Enabled = pCanEdit
    cmdUnSelectAll.Enabled = pCanEdit
    If pCanEdit = False AndAlso mvDocumentMenu IsNot Nothing Then
      mvDocumentMenu.DocumentNumber = 0
      dgr.ContextMenuStrip = Nothing
      dgr.SetToolBarVisible()
      Select Case mvExamDataType
        Case XMLExamDataSelectionTypes.ExamUnitLinkDocuments,
        XMLExamDataSelectionTypes.ExamCentreDocuments,
        XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments
          With dspTabGrid
            .DisplayGrid(0).ContextMenuStrip = Nothing
            .DisplayGrid(0).SetToolBarVisible()
            .DisplayGrid(1).ContextMenuStrip = Nothing
            .DisplayGrid(1).SetToolBarVisible()
            .DisplayGrid(2).ContextMenuStrip = Nothing
            .DisplayGrid(2).SetToolBarVisible()
          End With
      End Select
    End If
    If pCanEdit = True Then 'Careful, node edit-ability is set by Access Control.  We only override the edit-ability on top of the existing editability
      If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamCentreUnitLinkCategories Then
        'Exam centre unit categories grid displays links for the Unit and the centre unit.  The unit categories are for display only.  The user can create new categories but they will be centre-unit categories.
        If dgr.ActiveRow >= 0 AndAlso dgr.GetValue(dgr.ActiveRow, "ExamUnitLinkId").Length > 0 Then
          mvCurrentEpl.EnableControls(False, True)
          cmdDelete.Enabled = False
          cmdSave.Enabled = False
        End If
      End If
    End If
  End Sub

  Private Sub SetRowReadOnly()

    dgr.Refresh()
    For vRow As Integer = 0 To dgr.DataRowCount - 1
      If IntegerValue(dgr.GetValue(vRow, "ExamUnitLinkId")) > 0 Then
        dgr.SetCellsReadOnly(vRow, -1, True, True)
      End If
    Next
  End Sub

#End Region

#Region "Menu Handling"

  Private Sub mvEventMenu_ItemSelected(ByVal sender As Object, ByVal pMenuItem As ExamsMenu.ExamMenuItems) Handles mvExamsMenu.MenuSelected, mvSelExamsMenu.MenuSelected, mvExamsGridMenu.MenuSelected
    Dim vBusyCursor As New BusyCursor()
    Try
      Select Case pMenuItem
        Case ExamsMenu.ExamMenuItems.CreateProgramme
          Dim vFinder As New frmSimpleFinder
          vFinder.Init(CareNetServices.XMLLookupDataTypes.xldtExamSessions, True)
          If vFinder.ShowDialog() = DialogResult.OK Then
            Dim vList As New ParameterList(True)
            vList("ExamSessionCode") = vFinder.ResultValue
            vList.IntegerValue("ExamUnitId") = mvExamUnitId
            vList.IntegerValue("ExamUnitLinkId") = mvExamUnitLinkId

            Dim vParameters As ParameterList = ExamsDataHelper.CreateExamProgramme(vList)
            If vParameters.Contains("RowCount") Then
              ShowInformationMessage(InformationMessages.ImExamProgrammeCreated, vParameters.Item("RowCount"))
            End If
          End If

        Case ExamsMenu.ExamMenuItems.CopyLink
          Dim vExamInfo As ExamCopyInfo = New ExamCopyInfo(mvExamUnitId)
          Clipboard.Clear()
          Clipboard.SetData(GetType(ExamCopyInfo).FullName, vExamInfo)

        Case ExamsMenu.ExamMenuItems.PasteAsChild
          If Clipboard.ContainsData(GetType(ExamCopyInfo).FullName) Then
            Dim vExamInfo As ExamCopyInfo = DirectCast(Clipboard.GetData(GetType(ExamCopyInfo).FullName), ExamCopyInfo)
            Dim vList As New ParameterList(True)
            vList.IntegerValue("ExamUnitId1") = mvExamUnitId
            vList.IntegerValue("ExamUnitId2") = vExamInfo.ExamUnitId
            Dim vReturnList As ParameterList = ExamsDataHelper.AddItem(XMLExamDataSelectionTypes.ExamUnitLinks, vList)
            sel.Init(mvSelectorType)
            sel.SelectNode(mvExamUnitId, True)
          End If

        Case ExamsMenu.ExamMenuItems.Share
          If sel.SelectedNode IsNot Nothing AndAlso TypeOf sel.SelectedNode.Tag Is ExamSelectorItem Then
            Dim vSelectedItem As ExamSelectorItem = DirectCast(sel.SelectedNode.Tag, ExamSelectorItem)
            ShareItem(vSelectedItem)
          End If

        Case ExamsMenu.ExamMenuItems.RemoveLink
          Dim vList As New ParameterList(True)
          Dim vParentID As Integer = sel.GetParentID
          vList.IntegerValue("ExamUnitId1") = vParentID
          vList.IntegerValue("ExamUnitId2") = mvExamUnitId
          Dim vReturnList As ParameterList = ExamsDataHelper.DeleteItem(XMLExamDataSelectionTypes.ExamUnitLinks, vList)
          sel.Init(mvSelectorType)
          sel.SelectNode(vParentID, False)

        Case ExamsMenu.ExamMenuItems.Search
          If sender Is mvExamsMenu Then
            sel.DoSearchTree(Me)
          Else
            selExams.DoSearchTree(Me)
          End If

        Case ExamsMenu.ExamMenuItems.AddScheduleMultiple
          If mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule Then
            mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamSchedule
          Else
            mvExamMaintenanceType = XMLExamMaintenanceTypes.ExamMultipleSchedule
          End If
          RefreshCard()
        Case ExamsMenu.ExamMenuItems.Reallocate
          ' Ask user to select papers to reallocate
          Dim vList As New ParameterList(True)
          GetPrimaryKeyValues(vList, mvSelectedRow, False)
          vList.AddSystemColumns()
          Dim vDataSet As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, vList)
          Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
          If vDataTable IsNot Nothing Then
            If vDataTable.Rows.Count > 0 Then
              ' Stop the ContactName appearing as a hyper link in the Dialog based grid by renaming it
              If vDataTable.Columns.Contains("ContactName") Then
                vDataTable.Columns("ContactName").ColumnName = "CandidateName"
                DataHelper.GetColumnTableFromDataSet(vDataSet).Select("Name = 'ContactName'")(0).Item("Name") = "CandidateName"
              End If

              Dim vFrmSelect As New frmSelectItems(vDataSet, frmSelectItems.SelectItemsTypes.sitExamUnitMarkerAllocation)

              If vFrmSelect.ShowDialog() <> DialogResult.Cancel AndAlso vFrmSelect.SelectedValues.Length > 0 AndAlso vFrmSelect.SelectedNumbers.Length > 0 Then
                ' Ask user for target Marker
                vList.Clear()
                vList.AddConnectionData()
                vList.AddSystemColumns()
                vList.Add("ExamUnitId", dgr.GetValue(mvSelectedRow, "ExamUnitId"))
                vList.Add("ExamBookingUnitId", vFrmSelect.SelectedNumbers)
                Dim vExamPersonnelId As Integer = DisplayMarkerSelection(vList)
                If vExamPersonnelId > 0 Then

                  ' Ask user for target Marker Number
                  Dim vMarkerNumber As Integer = CInt(dgr.GetValue(mvSelectedRow, "MarkerNumber"))
                  If vMarkerNumber > 0 Then
                    vList.Remove("ExamPersonnelId")
                    vList.Remove("MarkerNumber")
                    vList.Add("ExamPersonnelId", vExamPersonnelId) 'New marker exam personnel id
                    vList.Add("MarkerNumber", vMarkerNumber) 'New Existing marker number
                    vList.Add("ExamMarkingBatchDetailId", vFrmSelect.SelectedValues)

                    ' Reallocate papers
                    If AllocatePapersToMarker(vList) Then RefreshCard()
                  End If
                End If
              End If
            End If
          End If
        Case ExamsMenu.ExamMenuItems.Unallocate
          Dim vList As New ParameterList(True)
          GetPrimaryKeyValues(vList, mvSelectedRow, False)
          vList.AddSystemColumns()
          Dim vDataSet As DataSet = ExamsDataHelper.GetExamData(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, vList)

          Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
          If Not vDataTable Is Nothing Then
            If vDataTable.Rows.Count > 0 Then
              ' Stop the ContactName appearing as a hyper link in the Dialog based grid by renaming it
              If vDataTable.Columns.Contains("ContactName") Then
                vDataTable.Columns("ContactName").ColumnName = "CandidateName"
                DataHelper.GetColumnTableFromDataSet(vDataSet).Select("Name = 'ContactName'")(0).Item("Name") = "CandidateName"
              End If

              Dim vFrmSelect As New frmSelectItems(vDataSet, frmSelectItems.SelectItemsTypes.sitExamUnitMarkerAllocation)

              If vFrmSelect.ShowDialog() <> DialogResult.Cancel AndAlso vFrmSelect.SelectedValues.Length > 0 Then
                vList.Clear()
                vList.AddConnectionData()
                vList.Add("ExamMarkingBatchDetailId", vFrmSelect.SelectedValues)
                ExamsDataHelper.DeleteItem(XMLExamDataSelectionTypes.ExamUnitMarkerAllocationList, vList)
                RefreshCard()
              End If
            End If
          End If
        Case ExamsMenu.ExamMenuItems.Clone
          ProcessClone()
      End Select

    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
        Case CareException.ErrorNumbers.enCannotUnshareExamUnitLinkedData, CareException.ErrorNumbers.enExamUnitMustHaveAtLeastOneParent
          ShowInformationMessage(vEx.Message)
        Case Else
          DataHelper.HandleException(vEx)
      End Select
    Catch vException As Exception
      If TypeOf (vException) Is NotSupportedException Then
        ShowInformationMessage(vException.Message)
      Else
        DataHelper.HandleException(vException)
      End If
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Private Enum MenuSource
    dgr
    dpl
    dgr0
    dgr1
    dgr2
    actiondgr0
    actiondgr4
    dplCustomise
    dplRevert
  End Enum

  Private Sub dgr0MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr0MenuNew.Click
    HandleMenuClick(False, XMLMaintenanceControlTypes.xmctDocumentTopic)
  End Sub
  Private Sub dgr1MenuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr1MenuNew.Click
    HandleMenuClick(False, XMLMaintenanceControlTypes.xmctDocumentLink)
  End Sub

  Private Sub dgr0MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr0MenuEdit.Click
    HandleMenuClick(True, XMLMaintenanceControlTypes.xmctDocumentTopic)
  End Sub

  Private Sub dgr1MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgr1MenuEdit.Click
    HandleMenuClick(True, XMLMaintenanceControlTypes.xmctDocumentLink)
  End Sub

  Private Sub HandleMenuClick(ByVal pEdit As Boolean, ByVal pMaintenanceType As XMLMaintenanceControlTypes)
    Dim vForm As frmCardMaintenance = Nothing
    Dim vCursor As New BusyCursor
    Dim vList As New ParameterList()
    Try

      Dim vRow As Integer = dspTabGrid.DisplayGrid(0).CurrentRow
      Dim vContactInfo As New ContactInfo(DataHelper.UserContactInfo.ContactNumber)
      vContactInfo.SelectedDocumentNumber = CInt(dgr.GetValue(dgr.ActiveRow, "DocumentNumber"))

      Select Case mvExamSelectorItem.ExamSelectionType
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCentreDocuments
          vList.Add("ExamCentreId", mvExamSelectorItem.ID)
          vList.Add("ExamCentre", "Y")
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnitLinkDocuments
          vList.Add("ExamUnitLinkId", mvExamSelectorItem.LinkID)
          vList.Add("ExamUnit", "Y")
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments
          vList.Add("ExamCentreUnitLinkId", mvExamSelectorItem.CentreUnitID)
          vList.Add("ExamCentreUnit", "Y")
      End Select

      Select Case pMaintenanceType
        Case XMLMaintenanceControlTypes.xmctDocumentTopic
          vForm = New frmCardMaintenance(Me, vContactInfo, XMLContactDataSelectionTypes.xcdtContactDocuments, mvDocAnalysisDataSource, pEdit, vRow, CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic, Nothing)
          ShowMaintenanceForm(vForm, CareServices.XMLMaintenanceControlTypes.xmctDocumentTopic)
        Case XMLMaintenanceControlTypes.xmctDocumentLink
          vForm = New frmCardMaintenance(Me, vContactInfo, XMLContactDataSelectionTypes.xcdtContactDocuments, mvDocLinkDataSource, pEdit, vRow, pMaintenanceType, vList)
          ShowMaintenanceForm(vForm, pMaintenanceType)
      End Select

    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub ShowMaintenanceForm(ByVal pForm As frmCardMaintenance, ByVal pMaintenanceType As CareServices.XMLMaintenanceControlTypes)
    If pForm IsNot Nothing Then
      mvCustomiseMenu = New CustomiseMenu
      Dim vMaintenanceType As CareServices.XMLMaintenanceControlTypes = pMaintenanceType
      If vMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctNone Then
        vMaintenanceType = pForm.MaintenanceType
      End If
      If vMaintenanceType <> CareNetServices.XMLMaintenanceControlTypes.xmctNone Then
        mvCustomiseMenu.SetContext(pForm, vMaintenanceType, "CON")
        pForm.SetCustomiseMenu(mvCustomiseMenu)
        Dim vLocation As Point = splRight.PointToScreen(splRight.Location)
        Dim vSize As Size = splRight.Size
        pForm.SetInitialBounds(vLocation, vSize)
        If MDIForm IsNot Nothing Then
          pForm.Show()
        Else
          pForm.Show(Me)
        End If
      End If
    End If
  End Sub
#End Region

#Region "Customisation"

  Private Sub CustomiseButtons(ByVal sender As Object, ByVal pDataSelectionType As Integer, ByVal pRevert As Boolean)
    Try
      Dim vCustomised As Boolean = DataHelper.CustomiseDisplayList(Me, pDataSelectionType, pRevert)
      If vCustomised Then
        'Image Buttons have been customised - need to close Form otherwise buttons not displayed correctly
        Me.Close()
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub mvCustomiseMenu_UpdatePanel(ByVal pRevert As Boolean) Handles mvCustomiseMenu.UpdatePanel
    RefreshCard()
    SetNodeDataEditable(sel.SelectedNode)
  End Sub

  Private Sub CustomiseExamSelector(ByVal sender As Object, ByVal pRevert As Boolean) Handles sel.TreeCustomised
    'ExamSelector tree nodes have been customised
    RefreshData()
  End Sub

#End Region

  Private Sub ShareItem(vSelectedItem As ExamSelectorItem)
    Dim vParamList As New ParameterList
    'vParamList("ExamUnitLinkId") = vList("Appeal2")
    vParamList.Add("SelectionType", ExamSelector.SelectionType.CourseSelection)
    vParamList.Add("RestrictionID", 0)
    vParamList.Add("RestrictionID2", vSelectedItem.UnitID)
    vParamList.Add("RestrictionID3", vSelectedItem.LinkID)
    vParamList.Add("ExamUnitDescription", FindControl(mvCurrentEpl, "ExamUnitDescription", False).Text)
    vParamList.Add("SessionID", vSelectedItem.SessionID)
    vParamList.Add("InitForSharingUnitID", vSelectedItem.UnitID)

    Dim vReturnParamList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptShareExamUnit, vParamList)
    If vReturnParamList IsNot Nothing AndAlso vReturnParamList.ContainsKey("SelectedExams") Then
      Dim vSelectedExams As List(Of ChangedItem) = TryCast(vReturnParamList.ObjectValue("SelectedExams"), List(Of ChangedItem))
      If vSelectedExams IsNot Nothing AndAlso vSelectedExams.Count > 0 Then
        Dim vCreateLinks As String = ",", vDeleteLinks As String = ","
        vSelectedExams.Where(Function(vChangedItem) vChangedItem.Checked).ToList().ForEach(Sub(vCreateItem) vCreateLinks &= vCreateItem.Item.LinkID & ",")
        vCreateLinks = vCreateLinks.Trim(",".ToCharArray())

        vSelectedExams.Where(Function(vChangedItem) vChangedItem.Checked = False).ToList().ForEach(Sub(vRemoveItem) vDeleteLinks &= vRemoveItem.Item.LinkID & ",")
        vDeleteLinks = vDeleteLinks.Trim(",".ToCharArray())

        Dim vWebParams As New ParameterList(True)
        vWebParams.IntegerValue("ExamUnitId2") = vSelectedItem.UnitID  'the highlighted unit in the Exams Maintenance is the child
        vWebParams.IntegerValue("SourceUnitLinkId") = vSelectedItem.LinkID ' the unit link id of the original child unit.  All the children of this item at that location (the grand-children) will also be copied
        vWebParams.Add("AddUnitLinks", vCreateLinks) 'the highlighted unit in the Exams Maintenance is the child
        vWebParams.Add("RemoveUnitLinks", vDeleteLinks)  'the highlighted unit in the Exams Maintenance is the child
        ExamsDataHelper.ShareExamUnitLink(vWebParams)
        RefreshData()
      End If
    End If
  End Sub

  Private Sub HideControlsByGradingMethod()

    If ExamsDataHelper.GradingMethod <> ExamGradingMethod.NG Then
      Dim vHideControls As New List(Of String)
      vHideControls.Add("IsGradingEndpoint")
      vHideControls.Add("ResultsReleaseDate")

      'Hide Controls - Consider extracting this to a method if it gets complicated
      For Each vItem As String In vHideControls
        If FindControl(mvCurrentEpl, vItem, False) IsNot Nothing Then mvCurrentEpl.SetControlVisible(vItem, False)
      Next
    End If
  End Sub

  Private Sub ResizePanels()
    If mvPanelProportions IsNot Nothing AndAlso mvPanelProportions.ContainsKey(mvExamDataType) Then splRight.SplitterDistance = CInt(splRight.Height * mvPanelProportions(mvExamDataType))
  End Sub


  Private Sub frmExams_Load(sender As Object, e As EventArgs) Handles Me.Load
    mvPanelProportions = New Dictionary(Of XMLExamDataSelectionTypes, Double)
    Dim vDataSelectionTypes As Array = [Enum].GetValues(GetType(XMLExamDataSelectionTypes))
    For Each vItem As XMLExamDataSelectionTypes In vDataSelectionTypes
      mvPanelProportions.Add(vItem, 0.5)
    Next
    mvPanelProportions(XMLExamDataSelectionTypes.ExamUnits) = 0.7 '0.7 is arbitrarily chosen.  Ideally should ask the panel what size it needs to be to display all its controls.  For future improvement.
    mvPanelProportions(XMLExamDataSelectionTypes.ExamCentres) = 0.7 'Setting to anything mopre than 70% somehow doesn't display the whole bottom panel with name change history
    mvPanelProportions(XMLExamDataSelectionTypes.ExamCentreDocuments) = 0.4
    mvPanelProportions(XMLExamDataSelectionTypes.ExamCentreUnitLinkDocuments) = 0.4
    mvPanelProportions(XMLExamDataSelectionTypes.ExamUnitLinkDocuments) = 0.4
    ResizePanels()
  End Sub

  Private Sub splRight_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles splRight.SplitterMoved
    Dim vProportion As Double = splRight.Panel1.Height / splRight.Height
    If Not mvPanelProportions.ContainsKey(mvExamDataType) Then
      mvPanelProportions.Add(mvExamDataType, vProportion)
    Else
      mvPanelProportions(mvExamDataType) = vProportion
    End If
  End Sub

  Public Sub SelectExamNode(ByVal pId As Integer, ByVal pDataSelected As XMLExamDataSelectionTypes, ByVal pType As String)
    Select Case pType
      Case "U"
        ProcessSelection(ExamSelector.SelectionType.Courses)
      Case "X"
        ProcessSelection(ExamSelector.SelectionType.Centres)
      Case "N"
        ProcessSelection(ExamSelector.SelectionType.Centres)
    End Select
    sel.SelectNode(pId, True, False, pDataSelected)
  End Sub

  Private Sub dgr_DocumentSelected(sender As Object, pRow As Integer, pDocumentNumber As Integer) Handles dgr.DocumentSelected
    Try
      FormHelper.EditDocument(pDocumentNumber, Me, Nothing)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub pnlCustomFormPage_EntitySelected(sender As Object, pEntityID As String, pEntityType As HistoryEntityTypes) Handles pnlCustomFormPage.EntitySelected
    Dim vEntityNumber As Integer
    If Integer.TryParse(pEntityID, vEntityNumber) Then
      MainHelper.NavigateHistoryItem(pEntityType, vEntityNumber, True)
    End If
  End Sub

  Private mvIsCloning As Boolean = False
  Private Property IsCloning As Boolean
    Get
      Return mvIsCloning
    End Get
    Set(value As Boolean)
      mvIsCloning = value
    End Set
  End Property

  Private Sub dgr_CPDCyclePeriodSelected(sender As Object, pRow As Integer, pCPDPeriodNumber As Integer) Handles dgr.CPDCyclePeriodSelected
    Dim vContactNumber As Integer = 0
    Dim vList As New ParameterList(True, True)
    vList.IntegerValue("ContactCpdPeriodNumber") = pCPDPeriodNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftCPDCyclePeriodFinder, vList))
    If vTable IsNot Nothing Then
      If vTable.Columns.Contains("ContactNumber") Then vContactNumber = IntegerValue(vTable.Rows(0).Item("ContactNumber").ToString)
    End If
    If vContactNumber > 0 Then FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPD, vContactNumber, False)
  End Sub

  Private Sub dgr_CPDPointSelected(sender As Object, pRow As Integer, pCPDPointNumber As Integer) Handles dgr.CPDPointSelected
    Dim vContactNumber As Integer = 0
    Dim vGotCyclePeriod As Boolean = False

    Dim vList As New ParameterList(True, True)
    vList.IntegerValue("ContactCpdPointNumber") = pCPDPointNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftCPDPointFinder, vList))
    If vTable IsNot Nothing Then
      If vTable.Columns.Contains("ContactNumber") Then vContactNumber = IntegerValue(vTable.Rows(0).Item("ContactNumber").ToString)
      If vTable.Columns.Contains("ContactCpdPeriodNumber") AndAlso IntegerValue(vTable.Rows(0).Item("ContactCpdPeriodNumber").ToString) > 0 Then vGotCyclePeriod = True
    End If

    Dim vType As CareNetServices.XMLContactDataSelectionTypes = CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPointsWithoutCycle
    If vGotCyclePeriod Then vType = CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPD
    If vContactNumber > 0 Then FormHelper.ShowCardIndex(vType, vContactNumber, False)
  End Sub

  Private Sub epl_ValidateItem(sender As Object, pParameterName As String, pValue As String, ByRef vValid As Boolean) Handles mvCurrentEpl.ValidateItem
    Select Case mvExamMaintenanceType
      Case XMLExamMaintenanceTypes.ExamUnit
        Select Case pParameterName
          Case "ExamUnitType"
            mvCurrentEpl.SetErrorField("ExamUnitType", "")
            If mvCourseSessionId > 0 Then
              'The Exam Unit Type must be a Question for session-based exams
              Dim vCtl As Control = mvCurrentEpl.FindPanelControl("ExamUnitType", False)
              If vCtl IsNot Nothing AndAlso TypeOf vCtl Is TextLookupBox Then
                Dim vTlb As TextLookupBox = DirectCast(vCtl, TextLookupBox)
                If vTlb.GetDataRowItem("ExamQuestion") <> "Y" Then 'Ignores the passed value and directly queries the selected value
                  mvCurrentEpl.SetErrorField("ExamUnitType", GetInformationMessage(InformationMessages.ImExamUnitTypeMustBeQuestion))
                  vValid = False
                End If
              End If
            End If
        End Select
    End Select
  End Sub
End Class
