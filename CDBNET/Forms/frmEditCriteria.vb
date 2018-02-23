Public Class frmEditCriteria

#Region "Event Declaration"
  Public Event ProcessSelection(ByVal pRunPhase As String, ByVal pParams As ParameterList)
  Public Event ProcessMailingCriteria(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByRef pSuccess As Boolean)
  Public Event ProcessMailingCriteriaWithOptional(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSet As Integer, ByVal pProcessVariables As Boolean, ByVal pEditSegmentCriteria As Boolean, ByRef pList As ParameterList, ByRef pSuccess As Boolean)
  Public Event CheckModule()
  Public Event SaveCriteria()
#End Region

#Region "Constants"
  Const TOKEN_END As Integer = 1
  Const TOKEN_COMMA As Integer = 2
  Const TOKEN_TO As Integer = 3
#End Region

#Region "Private Members"
  Private mvCriteriaDesc As String
  Private mvMailingTypeCode As String
  Private mvSelectionManager As Boolean
  Private mvFormUnLoad As Boolean
  Private mvMailingInfo As MailingInfo
  Private mvCriteriaSetDesc As String
  Private mvAllowSegmentOrgSelections As Boolean
  Private mvIncludeExclusions As Boolean
  Private mvDataSet As DataSet
  Private mvDataSetDefault As DataSet
  Private mvLicenceSeconds As Integer
  Private mvCurrentSS As DisplayGrid
  Private mvSelectionSet As Integer
  Private mvAddBrackets As Boolean
  Private mvCriteriaGridVisible As Boolean
  Private mvDefaultGridVisible As Boolean
#End Region

#Region "Property"
  Public ReadOnly Property SelectionManager() As Boolean
    Get
      Return mvSelectionManager
    End Get
  End Property

  Public ReadOnly Property Unload() As Boolean
    Get
      Return mvFormUnLoad
    End Get
  End Property

  Public WriteOnly Property SelectionSet() As Integer
    Set(ByVal value As Integer)
      mvSelectionSet = value
    End Set
  End Property
#End Region

  Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

  Public Sub New(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSetDesc As String, Optional ByVal pAllowSegmentOrgSelections As Boolean = False)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pMailingSelection, pCriteriaSetDesc, pAllowSegmentOrgSelections)
  End Sub

  Private Sub InitialiseControls(ByVal pMailingSelection As MailingInfo, ByVal pCriteriaSetDesc As String, Optional ByVal pAllowSegmentOrgSelections As Boolean = False)
    Dim vColWidth As String
    Dim vEditOnly As Boolean
    Dim vList As New ParameterList(True)
    mvMailingInfo = pMailingSelection
    mvCriteriaSetDesc = pCriteriaSetDesc
    mvDataSet = New DataSet
    SetControlTheme()

    If mvMailingInfo.AppealMailing _
      OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyPerformanceAnalyser _
      OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyScoringAnalyser _
      OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandardExclusions Then
      mvSelectionManager = False
      If mvMailingInfo.AppealMailing Then mvAllowSegmentOrgSelections = pAllowSegmentOrgSelections
    Else
      mvSelectionManager = True
    End If

    vColWidth = ","

    vEditOnly = Not mvSelectionManager

    Me.Text = String.Format(ControlText.FrmEditCriteria, mvMailingInfo.Caption) ' Criteria Selector

    dgrDefault.Visible = False
    dgrDefault.Enabled = False
    mvDefaultGridVisible = False
    vseStdExclOptions.Visible = False
    mvIncludeExclusions = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_include_exclusions, True)

    mvCurrentSS = dgrCriteria

    cmdOK.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdCount.Enabled = False
    ClearDesc()
    mvCriteriaDesc = Nothing

    With dgrCriteria
      .MaxGridRows = DisplayTheme.DefaultMaxGridRows
      .SetCellsEditable()
      .SetCellsReadOnly(-1, -1, True, True)
      .AllowRowMove()
      .AutoSetRowHeight = True
    End With

    With dgrDefault
      .MaxGridRows = DisplayTheme.DefaultMaxGridRows
      .SetCellsReadOnly(-1, -1, True, True)
      .AllowRowMove()
      .AutoSetRowHeight = True
    End With

    If vEditOnly Then
      vseInfo.Visible = False
      pnlTop.Visible = False
      cmdLists.Visible = False
      cmdSaveCriteria.Visible = False
      mvMailingInfo.NewCriteriaSet = mvMailingInfo.CriteriaSet
      GetCriteriaSetDetails()
      vseOptions.Visible = False
      pnlBottom.Visible = False
    Else
      GetDefaultCriteria()
      chkSkipCriteriaCount.Checked = True
    End If
    ChangeCurrentRow(mvCurrentSS, 0)
  End Sub

  ' TODO: Key Down event is not firing
  Private Sub dgrCriteria_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles dgrCriteria.KeyDown
    Try
      Dim vValue As Integer
      Dim vKeyCode As Integer

      If e.KeyCode = Keys.Shift Then
        If e.KeyCode = Keys.Up Then
          vKeyCode = 0
          vValue = -1
        ElseIf e.KeyCode = Keys.Down Then
          vKeyCode = 0
          vValue = 1
        End If

        With dgrCriteria
          'only allow row movement if your on a populated row
          If .CurrentDataRow < .RowCount Then
            'only allow the row to be moved up to row 1 or down to the last populated row
            If (vValue = -1 And .CurrentDataRow > 1) Or (vValue = 1 And .CurrentDataRow < (.RowCount - 1)) Then
              ' TODO: not sure to write this code
            End If
          End If
        End With
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrCriteria_GotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles dgrCriteria.GotFocus
    Try
      mvCurrentSS = dgrCriteria
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub dgrCriteria_LeaveCell(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer) Handles dgrCriteria.RowSelected
    Try
      ChangeCurrentRow(dgrCriteria, pRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrDefault_GotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles dgrDefault.GotFocus
    Try
      mvCurrentSS = dgrDefault
      ChangeCurrentRow(mvCurrentSS, mvCurrentSS.CurrentDataRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub dgrDefault_LeaveCell(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer) Handles dgrDefault.RowSelected
    Try
      ChangeCurrentRow(dgrDefault, pRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ChangeCurrentRow(ByVal vDgr As CDBNETCL.DisplayGrid, ByVal pNewRow As Integer)
    If GridRowExists(vDgr) Then
      If pNewRow > -1 Then
        With vDgr
          If vDgr Is dgrCriteria Then
            SetRowColor(vDgr)
            SetRowColor(vDgr, pNewRow)
          End If
          cmdDelete.Enabled = .GetValue(pNewRow, "IE").Length > 0
          cmdUpdate.Enabled = .GetValue(pNewRow, "IE").Length > 0
          If Not mvSelectionManager Then cmdCount.Enabled = .GetValue(pNewRow, "IE").Length > 0
        End With
      End If
    Else
      cmdDelete.Enabled = False
      cmdUpdate.Enabled = False
      If Not mvSelectionManager Then cmdCount.Enabled = False
    End If
    SetOK()
  End Sub

  Private Sub SetRowColor(ByVal vDgr As CDBNETCL.DisplayGrid, Optional ByVal pNewRow As Integer = -1)
    With vDgr
      If pNewRow > -1 Then
        .SelectRow(pNewRow, True)
        .BackColor = Color.Black      'black
        .ForeColor = Color.White      ' white
      Else  'colours defined by a preference setting
        'SetSSBackColor(pSS)
      End If
      .Refresh()
    End With
  End Sub

  Private Sub GetDefaultCriteria()

    mvCriteriaGridVisible = True
    Dim vList As New ParameterList(True)
    vList("CriteriaSet") = "0"

    ' Get dummy dataset to get columns which will be helpful to bind dataset to grid
    mvDataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstCriteriaSetDetails, CareServices.XMLTableDataSelectionTypes), vList)
    If Not mvDataSet.Tables.Contains("DataRow") Then
      AddDataRowTableToDataSet(mvDataSet)
    End If

    If Not AppValues.ConfigurationOption(AppValues.ConfigurationOptions.mail_pending_members, True) _
     AndAlso (mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembers OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyMembershipCards) Then
      With dgrCriteria
        If Not mvDataSet.Tables.Contains("DataRow") Then AddDataRowTableToDataSet(mvDataSet)
        Dim vNewDataRow As DataRow = mvDataSet.Tables("DataRow").NewRow()
        vNewDataRow("IE") = "I"
        vNewDataRow("CO") = "C"
        vNewDataRow("SearchArea") = "joined"
        vNewDataRow("Period") = DateAdd(DateInterval.Year, -100, Date.Today) & " to " & Date.Today.ToString
        mvDataSet.Tables("DataRow").Rows.Add(vNewDataRow)
      End With
    End If
    dgrCriteria.Populate(mvDataSet, False, True)
    SetupCols(dgrCriteria)

    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandingOrderCancellation OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGAYECancellation Then
      With dgrCriteria
        If Not mvDataSet.Tables.Contains("DataRow") Then AddDataRowTableToDataSet(mvDataSet)
        Dim vNewDataRow As DataRow = mvDataSet.Tables("DataRow").NewRow()
        vNewDataRow("IE") = "I"
        vNewDataRow("CO") = "C"
        vNewDataRow("SearchArea") = "cancreason"
        vNewDataRow("MainValue") = "NULL"
        mvDataSet.Tables("DataRow").Rows.Add(vNewDataRow)
        .Populate(mvDataSet, False, True)
        SetupCols(dgrCriteria)
      End With
    Else
      If vList.Contains("CriteriaSet") Then vList.Remove("CriteriaSet")
      vList("MarketingControls") = "Y"    ' dummy value which will help webservice to get details from marketing_controls table
      mvDataSetDefault = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstCriteriaSetDetails, CareServices.XMLTableDataSelectionTypes), vList)
      If mvDataSetDefault IsNot Nothing Then
        If Not mvDataSetDefault.Tables.Contains("DataRow") Then AddDataRowTableToDataSet(mvDataSetDefault)
      End If
      With dgrDefault
        If mvDataSetDefault.Tables.Contains("Column") AndAlso mvDataSetDefault.Tables("Column").Rows.Count > 0 Then
          For vIndex As Integer = 0 To mvDataSetDefault.Tables("Column").Rows.Count - 1
            If mvDataSetDefault.Tables("Column").Rows(vIndex)("Name").ToString = "AndOr" OrElse mvDataSetDefault.Tables("Column").Rows(vIndex)("Name").ToString = "LeftParenthesis" OrElse mvDataSetDefault.Tables("Column").Rows(vIndex)("Name").ToString = "RightParenthesis" OrElse mvDataSetDefault.Tables("Column").Rows(vIndex)("Name").ToString = "Counted" Then
              mvDataSetDefault.Tables("Column").Rows(vIndex)("Visible") = "N"
            End If
          Next
        End If
        .Populate(mvDataSetDefault, False, True)
        SetupCols(dgrDefault)

        If GridRowExists(dgrDefault) AndAlso .GetValue(0, "IE").Length > 0 Then
          .Visible = True
          .Enabled = True
          mvDefaultGridVisible = True
          For vRowIndex As Integer = 0 To .RowCount - 1
            If .GetValue(vRowIndex, "IE") = "I" Then
              .SetValue(vRowIndex, "I", "E")
              ShowInformationMessage(InformationMessages.ImStandardExcCriteria)  ' Standard Exclusion Criteria must all be Exclusions - Include Criteria will be modified
            End If
          Next
        End If
      End With
      ChangeCurrentRow(dgrDefault, 0)
      vseStdExclOptions.Visible = True
      vseStdExclOptions.Enabled = True
      optStandardExclusions2.Enabled = True
      optStandardExclusions.Enabled = True
      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGeneralMailing AndAlso mvIncludeExclusions Then
        optStandardExclusions.Checked = True
      Else
        optStandardExclusions2.Checked = True
      End If
    End If
  End Sub

  Private Sub GetCriteriaSetDetails()
    Dim vDesc As String = ""
    Dim vOwner As String = ""
    Dim vMerge As Boolean = False
    Dim vList As New ParameterList(True)
    vList.IntegerValue("CriteriaSet") = mvMailingInfo.NewCriteriaSet
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCriteriaSets, vList)
    If vTable IsNot Nothing Then
      If vTable.Rows.Count > 0 Then
        vDesc = vTable.Rows(0)("CriteriaSetDesc").ToString
        vOwner = vTable.Rows(0)("Owner").ToString
      End If
    End If

    With dgrCriteria
      Dim vRow As Integer = dgrCriteria.CurrentDataRow
      If vRow < 0 Then vRow = 0
      If GridRowExists(dgrCriteria) AndAlso dgrCriteria.GetValue(vRow, "IE").Length > 0 Then
        If ShowQuestion(QuestionMessages.QmReplaceCurrentCriteria, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then 'Replace current criteria?
          .Clear()
          mvDataSet.Tables("DataRow").Rows.Clear()
          mvCriteriaDesc = vDesc & "   (" & vOwner & ")   "
          SetCriteriaDesc(False)
        Else
          SetCriteriaDesc(True)
          vMerge = True
        End If
      Else
        mvCriteriaDesc = vDesc & "   (" & vOwner & ")   "
        SetCriteriaDesc(False)
      End If
      'Find first blank row
      .SelectRow(1, True)

      If vMerge Then
        If mvDataSet.Tables("DataRow") IsNot Nothing Then
          Dim vdsTemp As DataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstCriteriaSetDetails, CareServices.XMLTableDataSelectionTypes), vList)
          If vdsTemp IsNot Nothing AndAlso vdsTemp.Tables.Contains("DataRow") Then
            mvDataSet.Tables("DataRow").Merge(vdsTemp.Tables("DataRow"))
          End If
        End If
      Else
        mvDataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstCriteriaSetDetails, CareServices.XMLTableDataSelectionTypes), vList)
      End If
      If mvDataSet IsNot Nothing Then
        If mvDataSet.Tables.Contains("DataRow") Then
          For vRowIndex As Integer = 0 To mvDataSet.Tables("DataRow").Rows.Count - 1
            mvDataSet.Tables("DataRow").Rows(vRowIndex)("MainValue") = mvDataSet.Tables("DataRow").Rows(vRowIndex)("MainValue").ToString
            mvDataSet.Tables("DataRow").Rows(vRowIndex)("SubsidiaryValue") = mvDataSet.Tables("DataRow").Rows(vRowIndex)("SubsidiaryValue").ToString
            mvDataSet.Tables("DataRow").Rows(vRowIndex)("Period") = mvDataSet.Tables("DataRow").Rows(vRowIndex)("Period").ToString
          Next
        End If
      End If
      .Populate(mvDataSet, False, True)
      SetupCols(dgrCriteria)

      SafeSetFocus(dgrCriteria)
      ChangeCurrentRow(dgrCriteria, 0)
      .Visible = True
      mvCriteriaGridVisible = True
      dgrCriteria.Refresh()
    End With

    If mvSelectionManager Then
      If dgrCriteria.Visible AndAlso dgrCriteria.GetColumn("Counted") > 0 Then dgrCriteria.SetColumnVisible("Counted", False)
      If dgrDefault.Visible AndAlso dgrDefault.GetColumn("Counted") > 0 Then dgrDefault.SetColumnVisible("Counted", False)
    End If
  End Sub

  Private Sub SafeSetFocus(ByVal pControl As Control)
    If pControl.Enabled = True And pControl.Visible = True Then pControl.Focus()
  End Sub

  Private Sub SetupCols(ByVal pDgr As DisplayGrid)
    With pDgr
      Dim vItemsData() As String = {}

      vItemsData = Convert.ToString(",and,or").Split(CChar(","))
      .SetComboBoxColumn("AndOr", vItemsData, vItemsData)

      vItemsData = Convert.ToString(",(,((,(((,((((,(((((,((((((,(((((((,((((((((").Split(CChar(","))
      .SetComboBoxColumn("LeftParenthesis", vItemsData, vItemsData)

      vItemsData = Convert.ToString(",),)),))),)))),))))),)))))),))))))),))))))))").Split(CChar(","))
      .SetComboBoxColumn("RightParenthesis", vItemsData, vItemsData)

      If mvDataSet.Tables.Contains("Column") Then
        For Each vColumn As DataRow In mvDataSet.Tables("Column").Rows
          If GridRowExists(pDgr) Then .SetBackgroundColour(vColumn("Name").ToString, Color.White)
          If vColumn("Name").ToString = "AndOr" OrElse vColumn("Name").ToString = "LeftParenthesis" OrElse vColumn("Name").ToString = "RightParenthesis" Then .SetPreferredColumnWidth(.GetColumn(vColumn("Name").ToString))
        Next
      End If

      Dim vRowIndex As Integer = 0
      If pDgr Is dgrCriteria Then
        If mvDataSet.Tables.Contains("DataRow") Then
          For Each vRow As DataRow In mvDataSet.Tables("DataRow").Rows
            .SetValue(vRowIndex, "AndOr", vRow("AndOr").ToString)
            .SetValue(vRowIndex, "LeftParenthesis", vRow("LeftParenthesis").ToString)
            .SetValue(vRowIndex, "RightParenthesis", vRow("RightParenthesis").ToString)
            vRowIndex += 1
          Next
        End If
      Else
        If mvDataSetDefault.Tables.Contains("DataRow") Then
          For Each vRow As DataRow In mvDataSetDefault.Tables("DataRow").Rows
            .SetValue(vRowIndex, "AndOr", vRow("AndOr").ToString)
            .SetValue(vRowIndex, "LeftParenthesis", vRow("LeftParenthesis").ToString)
            .SetValue(vRowIndex, "RightParenthesis", vRow("RightParenthesis").ToString)
            vRowIndex += 1
          Next
        End If
      End If
      ShowGridCountColumn(pDgr)
      If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandardExclusions Then
        pDgr.SetColumnReadOnly("IE", True)
      End If
    End With
  End Sub

  Private Sub ShowGridCountColumn(ByVal pDgr As DisplayGrid)
    If mvSelectionManager Then
      If pDgr.Visible Then pDgr.SetColumnVisible("Counted", False)
    End If
  End Sub

  Private Function GridRowExists(ByVal pDgr As DisplayGrid) As Boolean
    Return pDgr.RowCount > 0
  End Function

  Private Sub SetCriteriaDesc(ByVal pAmended As Boolean)
    Dim vCurrent As String

    If mvCriteriaDesc IsNot Nothing AndAlso mvCriteriaDesc.Length > 0 Then
      vCurrent = mvCriteriaDesc
    Else
      vCurrent = ControlText.LblCurrent   ' Current
      pAmended = False
    End If
    If pAmended Then
      lblCurrentCriteria.Text = String.Format(ControlText.LblCurrentCriteriaAmmended, vCurrent)    'Selected Criteria: {0} Amended
    Else
      lblCurrentCriteria.Text = String.Format(ControlText.LblCurrentCriteriaSelected, vCurrent)    'Selected Criteria: {0} 
    End If
  End Sub

  Private Sub ClearDesc()
    lblCurrentList.Text = ControlText.LblCurrentNoList          'No List Selected
    lblCurrentCriteria.Text = ControlText.LblNoCriteriaSelected      'No Criteria Selected
  End Sub

  Private Sub frmEditCriteria_FormClosing_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    Try
      mvFormUnLoad = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      'We will assume that this can only occur if There is a criteria set
      'or there is a selection set or both
      'Save the current criteria set
      mvAddBrackets = False
      If PrepareMailing() Then
        If mvSelectionManager Then
          Me.Enabled = False
          mvMailingInfo.BypassCriteriaCount = chkSkipCriteriaCount.Checked
          ProcessMailingSelection()
          Me.Enabled = True
          Select Case mvMailingInfo.GenerateStatus
            Case MailingInfo.MailingGenerateResult.mgrRefine
              ClearCriteria()
              lblCurrentList.Text = String.Format(ControlText.LblCurrentSelectedList2, AppValues.Logname, mvMailingInfo.SelectionCount) ' Selected List: Current Set   ({0})   {1} Records
            Case MailingInfo.MailingGenerateResult.mgrReset
              ClearAll()
          End Select
        Else
          If mvAllowSegmentOrgSelections Then RaiseEvent ProcessMailingCriteriaWithOptional(mvMailingInfo, mvMailingInfo.CriteriaSet, False, True, Nothing, True) 'ProcessMailingCriteria(mvMailingInfo.CriteriaSet, False, True)
          Me.Close()
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
      Me.Enabled = True
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCriteria_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCriteria.Click
    Try
      Dim vForm As frmCriteriaLists
      mvMailingInfo.NewCriteriaSet = mvMailingInfo.CriteriaSet
      vForm = New frmCriteriaLists(mvMailingInfo)
      vForm.ShowDialog()
      If mvMailingInfo.NewCriteriaSet > 0 Then GetCriteriaSetDetails()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdLists_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLists.Click
    Try
      mvMailingInfo.NewSelectionSet = 0
      Dim vForm As New frmGenMLists(mvMailingInfo.MailingTypeCode, mvMailingInfo) 'BR17264 - Deal with frmGenMLists directly, consistant with Criteria 
      vForm.ShowDialog()
      If mvMailingInfo.NewSelectionSet > 0 Then
        GetMailingSelectionSet()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    Try
      Dim vResult As Integer = 0
      Dim vForm As frmAddCriteria
      If mvCurrentSS Is dgrCriteria Then RemoveRows()
      SetRowColor(dgrCriteria)
      mvCurrentSS.SelectRow(mvCurrentSS.RowCount - 1, True)
      vForm = New frmAddCriteria(mvMailingInfo, (mvCurrentSS Is dgrDefault) Or (mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandardExclusions))
      mvMailingInfo.CurrentCriteria.Valid = False
      vForm.ShowDialog()
      If mvMailingInfo.CurrentCriteria.Valid Then SetFromCurrentCriteria(True)
      ChangeCurrentRow(mvCurrentSS, mvCurrentSS.CurrentDataRow)

      If mvCurrentSS Is dgrDefault Then
        With mvCurrentSS
          If GridRowExists(mvCurrentSS) Then
            If vseStdExclOptions.Visible And Not vseStdExclOptions.Enabled And .GetValue(0, "IE").Length > 0 Then
              vseStdExclOptions.Enabled = True
              optStandardExclusions2.Enabled = True
              optStandardExclusions.Enabled = True
              If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGeneralMailing And mvIncludeExclusions Then
                optStandardExclusions.Checked = True
              Else
                optStandardExclusions2.Checked = True
              End If
            End If
          End If
        End With
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
    Try
      UpdateData()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      If mvCurrentSS.CurrentDataRow = -1 Then
        ShowInformationMessage(InformationMessages.ImSelectRow)
        Exit Sub
      End If
      With mvCurrentSS
        If .CurrentDataRow >= 0 Then
          .DeleteRow(.CurrentDataRow)
          If mvDataSet IsNot Nothing Then mvDataSet.AcceptChanges()
        End If
        If .CurrentDataRow > 0 Then .SelectRow(.CurrentDataRow - 1, False)
        ChangeCurrentRow(mvCurrentSS, .CurrentDataRow)
        If mvCurrentSS Is dgrCriteria Then
          If GridRowExists(dgrCriteria) AndAlso .GetValue(0, "IE").Length = 0 Then
            mvCriteriaDesc = ""
            lblCurrentCriteria.Text = ControlText.LblNoCriteriaSelected  ' No Criteria Selected
          Else
            SetCriteriaDesc(True)
          End If
        Else
          vseStdExclOptions.Enabled = vseStdExclOptions.Visible And (GridRowExists(dgrDefault) AndAlso .GetValue(0, "IE").Length > 0)
          optStandardExclusions2.Enabled = vseStdExclOptions.Enabled
          optStandardExclusions.Enabled = vseStdExclOptions.Enabled
          If Not vseStdExclOptions.Enabled Then optStandardExclusions2.Checked = True
        End If
      End With
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSaveCriteria_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveCriteria.Click
    Try
      Dim vPreProcess As Boolean
      If ValidateBrackets(dgrCriteria) Then
        'We don't want to include std criteria, so force the pre-process option button to true.
        vPreProcess = optStandardExclusions2.Checked
        optStandardExclusions2.Checked = True
        mvMailingInfo.CriteriaRows = SaveCriteriaSetDetails(mvMailingInfo.CriteriaSet)
        optStandardExclusions2.Checked = vPreProcess
        optStandardExclusions.Checked = Not vPreProcess
        If mvMailingInfo.CriteriaRows = 0 Then
          RaiseEvent SaveCriteria()
        Else
          Dim vSuccess As Boolean = True
          RaiseEvent ProcessMailingCriteria(mvMailingInfo, mvMailingInfo.CriteriaSet, vSuccess)
          If vSuccess Then
            RaiseEvent SaveCriteria()
          End If
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ProcessMailingSelection()
    If mvMailingInfo.CriteriaRows = 0 Then
      RaiseEvent ProcessSelection("Phase1", Nothing)
    Else
      Dim vList As New ParameterList(True)
      Dim vSuccess As Boolean = True
      RaiseEvent ProcessMailingCriteriaWithOptional(mvMailingInfo, mvMailingInfo.CriteriaSet, True, False, vList, vSuccess)
      If vSuccess Then
        RaiseEvent ProcessSelection("Phase1", vList)
      End If
    End If
  End Sub

  Private Sub Count(ByVal pCriteriaSet As Integer, ByVal pList As ParameterList)
    Dim vSupressCount As Boolean = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.ma_suppress_max_count, True)
    Dim vContainsORs As Boolean = mvMailingInfo.CriteriaContainsORs(pCriteriaSet)
    Dim vMinCount As Integer
    Dim vResult As DialogResult
    Dim vStop As Boolean = False

    If vSupressCount = False Then
      'we only want to do this bit if config says we're doing the max count
      If vContainsORs Then
        vMinCount = 99999999
        vResult = ShowQuestion(QuestionMessages.QmPerformFullCount, MessageBoxButtons.YesNo)
      Else
        vMinCount = mvMailingInfo.GetMailingSelectionRoughCount(pCriteriaSet, pList)
      End If
      If vMinCount = 0 Then
        'no contacts found in initial search so don't do anything else
        ShowInformationMessage(InformationMessages.ImNoContactsFound)
        vStop = True
      End If
      'either contacts have been found, or we're not doing max count
      If Not vContainsORs Then
        If vMinCount > 0 Then
          ShowInformationMessage(String.Format(InformationMessages.ImCriteriaMaxCount, vMinCount.ToString))
        Else
          ShowInformationMessage(String.Format(InformationMessages.ImCriteriaQuickCount, Math.Abs(vMinCount).ToString))
        End If
        vResult = ShowQuestion(QuestionMessages.QmCriteriaPreciseAnswer, MessageBoxButtons.YesNo)
      End If
      'if no has been selected to a detailed count then we don't want to do any more
      If Not vStop AndAlso vResult = vbNo Then
        vStop = True
        mvMailingInfo.SelectionCount = Math.Abs(vMinCount)
      End If
    End If
    If Not vStop Then       'we are doing detailed count
      RaiseEvent ProcessSelection("Count", pList)
      vMinCount = mvMailingInfo.SelectionCount
      If vMinCount > 0 Then
        ShowInformationMessage(String.Format(InformationMessages.ImCriteriaContactFound, vMinCount.ToString))
      Else
        ShowInformationMessage(InformationMessages.ImCriteriaNoContactFound)
      End If
    End If
  End Sub

  Private Sub ClearAll()
    If ShowQuestion(QuestionMessages.QmDeleteCriteriaSelection, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then  'Delete criteria and selection?
      If mvSelectionManager AndAlso mvSelectionSet > 0 Then mvMailingInfo.DeleteSelectionSetMailing(mvSelectionSet, True) 'DataHelper.DeleteMailingSelectionSet(mvSelectionSet, mvMailingTypeCode, 0, True) 
      mvMailingInfo.Revision = 0
      ClearDesc()
      ClearCriteria()
      If mvSelectionManager Then GetDefaultCriteria()
      SafeSetFocus(dgrCriteria)
    End If
  End Sub

  Private Sub ClearCriteria()
    Dim vDelete As Boolean
    If GridRowExists(dgrCriteria) Then vDelete = True
    With dgrCriteria
      Dim vRowIndex As Integer = 0
      While vRowIndex < .RowCount
        .DeleteRow(vRowIndex)
        vRowIndex = vRowIndex + 1
      End While
    End With
    If mvDataSet IsNot Nothing Then mvDataSet.AcceptChanges()

    If vDelete Then
      If mvMailingInfo.CriteriaSet > 0 Then
        Dim vList As New ParameterList(True)
        vList.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
        DataHelper.DeleteCriteriaSetDetails(vList)
      End If
      mvCriteriaDesc = ""
      lblCurrentCriteria.Text = ControlText.LblNoCriteriaSelected ' No Criteria Selected
      SetOK()
    End If
  End Sub

  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Try
      ClearAll()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click
    Try
      If mvCurrentSS.CurrentDataRow = -1 Then
        ShowInformationMessage(InformationMessages.ImSelectRow)
        Exit Sub
      End If
      Dim vRows As Integer
      Dim vCriteriaSet As Integer
      Dim vCount As Integer = 0
      Dim vHoldCriteria As Integer
      mvAddBrackets = False
      If mvSelectionManager Then
        'Save the current criteria set
        vRows = SaveCriteriaSetDetails(mvMailingInfo.CriteriaSet)
      Else
        vHoldCriteria = mvMailingInfo.CriteriaSet
        vCriteriaSet = 0
        vRows = SaveCriteriaSetDetails(vCriteriaSet, False, False)
        mvMailingInfo.CriteriaSet = vCriteriaSet
      End If

      If PrepareMailing() Then ProcessMailingCount()

      If Not mvSelectionManager Then
        Dim vList As New ParameterList(True)
        vList.IntegerValue("CriteriaSet") = mvMailingInfo.CriteriaSet
        DataHelper.DeleteCriteriaSetDetails(vList)
        mvMailingInfo.CriteriaSet = vHoldCriteria
        With dgrCriteria
          If mvMailingInfo.SelectionCount > 0 Then .SetValue(.CurrentDataRow, "Counted", mvMailingInfo.SelectionCount.ToString)
        End With
      End If
    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enVarInMultipleAreas
          ShowInformationMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enInvalidCharAfter
          ShowErrorMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enVariableNameContainsInvalidCharacters
          ShowErrorMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ProcessMailingCount()
    Dim vList As New ParameterList(True)
    vList("Mail") = "N"
    Dim vSuccess As Boolean = True
    RaiseEvent ProcessMailingCriteriaWithOptional(mvMailingInfo, mvMailingInfo.CriteriaSet, True, False, vList, vSuccess)
    If vSuccess Then Count(mvMailingInfo.CriteriaSet, vList)
  End Sub

  Private Function PrepareMailing() As Boolean
    Dim vContinue As Boolean
    Dim vRemoveExclusionSet As Boolean
    vContinue = ValidateBrackets(dgrCriteria)
    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyIrishGiftAid Then vContinue = CheckMandatoryCriteria()
    If vContinue Then
      If optStandardExclusions.Checked Then CheckORs()
      mvMailingInfo.CriteriaRows = SaveCriteriaSetDetails(mvMailingInfo.CriteriaSet)
      If optStandardExclusions2.Checked Then
        With dgrDefault
          If GridRowExists(dgrDefault) AndAlso .GetValue(0, "IE").Length > 0 Then
            If mvMailingInfo.ExclusionCriteriaSet = 0 Then
              mvMailingInfo.ExclusionCriteriaSet = 0
            End If
            SaveCriteriaSetDetails(mvMailingInfo.ExclusionCriteriaSet, True, True)
          Else
            vRemoveExclusionSet = True
          End If
        End With
      Else
        vRemoveExclusionSet = True
      End If
      If vRemoveExclusionSet Then
        If mvMailingInfo.ExclusionCriteriaSet > 0 Then
          Dim vList As New ParameterList(True)
          vList.IntegerValue("CriteriaSet") = mvMailingInfo.ExclusionCriteriaSet
          DataHelper.DeleteCriteriaSetDetails(vList)
        End If
        mvMailingInfo.ExclusionCriteriaSet = 0
      End If
      vContinue = True
    End If
    Return vContinue
  End Function

  Private Sub CheckORs()
    'This will check whether there are any "or"'s in the criteria
    'And if there are any, the entire criteria and the exclusions will have brackets added
    'This is only called when the 'Include Exclusion ..' checkbox had been checked
    Dim vRow As Integer
    Dim vLastRow As Integer

    vRow = 1
    vLastRow = dgrCriteria.RowCount
    mvAddBrackets = False

    With dgrCriteria
      For vRow = 0 To vLastRow - 1
        If .GetValue(vRow, "AndOr").Length > 0 Then
          If .GetValue(vRow, "AndOr").ToLower() = "or" Then mvAddBrackets = True
        End If
        If mvAddBrackets Then Exit For
      Next
    End With
  End Sub

  Private Function CheckMandatoryCriteria() As Boolean
    Dim vLastRow As Integer
    Dim vRow As Integer
    Dim vValid As Boolean

    vRow = 1
    vLastRow = dgrCriteria.RowCount - 1
    vValid = False

    If mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyIrishGiftAid Then
      With dgrCriteria
        For vRow = 0 To vLastRow
          If .GetValue(vRow, "SearchArea").ToLower() = "value" Then vValid = True
          If vValid = True Then Exit For
        Next
      End With
    Else
      vValid = True
    End If
    If vValid = False Then ShowInformationMessage(InformationMessages.ImIrishGiftAidMailings) ' Irish Gift Aid Mailings require the 'Value of Payments' search area to be used.
    Return vValid
  End Function

  Private Sub DisplayGridTag()
    If dgrCriteria.Visible Then
      grpCurrentCriteria.Text = ControlText.LblCurrentCriteria ' Current Criteria
    ElseIf dgrDefault.Visible Then
      grpStandardExclusions.Text = ControlText.LblStandardExclusions ' Standard Exclusions
    Else
      lblCurrentCriteria.Text = String.Empty
    End If
  End Sub

  Private Sub dgr_RowDoubleClicked(ByVal sender As System.Object, ByVal pRow As System.Int32) Handles dgrCriteria.RowDoubleClicked
    Try
      UpdateData()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub UpdateData()
    If mvCurrentSS.CurrentDataRow = -1 Then
      ShowInformationMessage(InformationMessages.ImSelectRow)
      Exit Sub
    End If
    Dim vCD As CriteriaDetails
    vCD = mvMailingInfo.CurrentCriteria
    With mvCurrentSS
      vCD.IE = .GetValue(.CurrentDataRow, "IE")
      vCD.CO = .GetValue(.CurrentDataRow, "CO")
      vCD.SearchArea = .GetValue(.CurrentDataRow, "SearchArea")
      vCD.MainValue = ParseDetail(.GetValue(.CurrentDataRow, "MainValue"))
      vCD.SubsidiaryValue = .GetValue(.CurrentDataRow, "SubsidiaryValue")
      vCD.Period = .GetValue(.CurrentDataRow, "Period")
      vCD.Valid = True
    End With
    Dim vForm As New frmAddCriteria(mvMailingInfo, (mvCurrentSS Is dgrDefault) Or (mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandardExclusions))
    vForm.ShowDialog()
    If mvMailingInfo.CurrentCriteria.Valid Then SetFromCurrentCriteria()
  End Sub

  Private Sub timPolling_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timPolling.Tick
    Try
      If mvMailingTypeCode = "GM" Then
        mvLicenceSeconds = CInt(mvLicenceSeconds + (timPolling.Interval / 1000))
        If mvLicenceSeconds >= 30 * 60 Then    '30 minutes for now
          RaiseEvent CheckModule()
          mvLicenceSeconds = 0
        End If
      End If

    Catch ex As Exception

    End Try
  End Sub

  Private Sub GetMailingSelectionSet()
    Dim vResponse As Boolean
    Dim vDesc As String = String.Empty
    Dim vOwner As String = String.Empty
    Dim vCount As Integer

    If mvMailingInfo.Revision = 0 Then
      vResponse = True
    Else
      vResponse = ShowQuestion(QuestionMessages.QmReplaceCurrentSelection, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes
      If vResponse Then mvMailingInfo.DeleteSelection(mvMailingInfo.SelectionSet, 0)
    End If
    If vResponse Then
      Dim vList As New ParameterList(True)
      vList("ApplicationName") = mvMailingInfo.MailingTypeCode
      If mvMailingInfo.SelectionSet > 0 Then vList.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
      vList.IntegerValue("NewSelectionSetNumber") = mvMailingInfo.NewSelectionSet
      vList.IntegerValue("OriginalRevision") = 0
      vList.IntegerValue("NewRevision") = 1
      vList("Unique") = "N"
      vList("TempToHold") = "N"
      DataHelper.UpdateMailingSelectionSet(vList)
      mvMailingInfo.Revision = 1
      ''Set the name of the current list
      vList.Clear()
      vList = New ParameterList(True)
      vList.IntegerValue("SelectionSetNumber") = mvMailingInfo.NewSelectionSet
      Dim vDataSet As DataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstSelectionSetData, CareServices.XMLTableDataSelectionTypes), vList)
      If vDataSet IsNot Nothing Then
        If vDataSet.Tables.Contains("DataRow") Then
          If vDataSet.Tables("DataRow").Rows.Count > 0 Then
            vDesc = vDataSet.Tables("DataRow").Rows(0)("SelectionSetDesc").ToString
            vCount = CInt(vDataSet.Tables("DataRow").Rows(0)("NumberInSet").ToString)
            vOwner = vDataSet.Tables("DataRow").Rows(0)("UserName").ToString
          End If
        End If
      End If
      lblCurrentList.Text = String.Format(ControlText.LblCurrentSelectedList, vDesc, vOwner, vCount) ' Selected List: {0}   ({1})  {2} Records
      SetOK()
    End If
  End Sub

  Private Function EndToken(ByVal pSource As String, ByRef pLength As Integer) As Integer
    'Search a string for commas or the word 'to'
    'Return a token for comma, to or neither
    Dim vS As String

    EndToken = TOKEN_END
    For pLength = 1 To pSource.Length
      vS = Strings.Mid(pSource, pLength, 1)
      If vS = "," Then
        Return TOKEN_COMMA
      ElseIf vS = " " And Strings.Mid(pSource, pLength, 4) = " to " Then
        Return TOKEN_TO
      End If
    Next
  End Function

  Private Function StripToken(ByVal pSource As String, ByVal pPosition As Integer, ByVal pLength As Integer) As String
    'Strip a token out of the given string - Remove bounding double quotes if found
    Dim vString As String
    vString = Strings.Mid(pSource, pPosition, pLength).Trim()
    If InStr(vString, " ") <> 0 Then    ' Do not remove double quotes if there is a space between
      'Do Nothing
    Else
      If Strings.Left(vString, 1) = """" And Strings.Right(vString, 1) = """" Then vString = Strings.Mid(vString, 2, Len(vString) - 2)
    End If
    Return vString
  End Function

  Private Function ParseDetail(ByVal pValue As String) As String
    'Parse out the different items in a value string
    'They will be separated by commas or by the TO keyword
    'Multiple items will be added to the result string with returns
    Dim vToken As Integer
    Dim vLength As Integer
    Dim vPosition As Integer
    Dim vStartVal As String = String.Empty
    Dim vResult As String = String.Empty

    vPosition = 1
    Do
      vToken = EndToken(Strings.Mid(pValue, vPosition), vLength)
      If vToken = TOKEN_COMMA Or vToken = TOKEN_END Then
        If vResult.Length > 0 Then vResult = vResult & Environment.NewLine
        If vStartVal <> String.Empty Then
          'Add 'Token to Token'
          vResult = vResult & vStartVal & " to " & StripToken(pValue, vPosition, vLength - 1)
        Else
          vResult = vResult & StripToken(pValue, vPosition, vLength - 1)
        End If
        vStartVal = String.Empty
        vPosition = vPosition + vLength
      ElseIf vToken = TOKEN_TO Then
        vStartVal = StripToken(pValue, vPosition, vLength - 1)
        vPosition = vPosition + (vLength + 3)
      End If
    Loop While vToken <> TOKEN_END
    Return vResult
  End Function

  Private Function SaveCriteriaSetDetails(ByRef pSetNumber As Integer, Optional ByVal pEntireSet As Boolean = True, Optional ByVal pExclusionSet As Boolean = False) As Integer
    'Save the criteria set details the user has been editing
    'Use the control_number we created at form load time
    Dim vCount As Integer
    Dim vSequenceNumber As Integer

    If pEntireSet AndAlso pSetNumber > 0 Then
      Dim vList As New ParameterList(True)
      vList.IntegerValue("CriteriaSet") = pSetNumber
      vList("DeleteAllCriteriaDetails") = "Y"
      DataHelper.DeleteCriteriaSetDetails(vList)
    End If

    vSequenceNumber = 1
    If Not pExclusionSet Then
      RemoveRows()
      vCount = SaveCriteriaSetDetailsSS(dgrCriteria, pSetNumber, vSequenceNumber, pEntireSet)
      If optStandardExclusions.Checked Then pExclusionSet = True
    End If
    If pExclusionSet Then vCount = vCount + SaveCriteriaSetDetailsSS(dgrDefault, pSetNumber, vSequenceNumber, pEntireSet)
    Return vCount    'Number of rows stored
  End Function

  Private Function SaveCriteriaSetDetailsSS(ByVal pDgr As DisplayGrid, ByRef pCriteriaSet As Integer, ByRef pSequenceNumber As Integer, ByVal pEntireSet As Boolean) As Integer
    Dim vCount As Integer
    Dim vStartRow As Integer
    Dim vEndRow As Integer
    Dim vContinue As Boolean

    If pEntireSet Then
      vStartRow = 0
    Else
      vStartRow = pDgr.CurrentDataRow
    End If
    vEndRow = pDgr.RowCount - 1
    vContinue = True

    With pDgr
      Dim vRowIndex As Integer = vStartRow
      If vRowIndex > -1 Then
        While vRowIndex <= vEndRow AndAlso .GetValue(vRowIndex, "IE").Length > 0 AndAlso vContinue
          Dim vList As New ParameterList(True)
          vList.IntegerValue("CriteriaSet") = pCriteriaSet
          vList("ApplicationName") = mvMailingInfo.MailingTypeCode
          vList.IntegerValue("SequenceNumber") = pSequenceNumber
          If vRowIndex = 0 Then .SetValue(vRowIndex, "AndOr", "")
          vList("AndOr") = .GetValue(vRowIndex, "AndOr").ToLower()

          If pDgr Is dgrDefault AndAlso optStandardExclusions.Checked AndAlso .GetValue(vRowIndex, "AndOr").Length > 0 Then
            vList("AndOr") = ""
          End If

          vList("LeftParenthesis") = .GetValue(vRowIndex, "LeftParenthesis")
          If (mvAddBrackets = True And vRowIndex = vStartRow) Then vList("LeftParenthesis") = vList("LeftParenthesis") & "("

          vList("IE") = .GetValue(vRowIndex, "IE")
          vList("CO") = .GetValue(vRowIndex, "CO")
          vList("SearchArea") = .GetValue(vRowIndex, "SearchArea")
          vList("MainValue") = .GetValue(vRowIndex, "MainValue").Replace(Environment.NewLine, ", ")
          vList("SubsidiaryValue") = .GetValue(vRowIndex, "SubsidiaryValue").Replace(Environment.NewLine, ", ")
          vList("Period") = .GetValue(vRowIndex, "Period").Replace(Environment.NewLine, ", ")
          vList("RightParenthesis") = .GetValue(vRowIndex, "RightParenthesis")
          If (mvAddBrackets = True And vRowIndex = vEndRow) Then vList("RightParenthesis") = vList("RightParenthesis") & ")"

          Dim vResult As ParameterList = DataHelper.AddCriteriaSetDetails(vList)

          If vResult IsNot Nothing Then
            pCriteriaSet = vResult.IntegerValue("CriteriaSet")
          End If

          If pEntireSet Then
            vRowIndex = vRowIndex + 1
            vCount = vCount + 1
            pSequenceNumber = pSequenceNumber + 1
          Else
            'Only save the active row
            vContinue = False
          End If
        End While
      End If
    End With
    Return vCount     'Number of rows stored
  End Function

  Private Sub SetFromCurrentCriteria(Optional ByVal pNewRow As Boolean = False)
    Dim vCD As CriteriaDetails
    Dim vRow As Integer

    If pNewRow Then
      vRow = mvCurrentSS.RowCount
    Else
      vRow = mvCurrentSS.CurrentDataRow
    End If
    vCD = mvMailingInfo.CurrentCriteria
    If vCD.Valid Then
      With mvCurrentSS
        If pNewRow Then
          If mvCurrentSS Is dgrCriteria Then
            If mvDataSet IsNot Nothing Then
              If Not mvDataSet.Tables.Contains("DataRow") Then AddDataRowTableToDataSet(mvDataSet)
              If mvDataSet.Tables.Contains("DataRow") Then
                Dim vDataRow As DataRow = mvDataSet.Tables("DataRow").NewRow()
                vDataRow("SearchArea") = vCD.SearchArea
                vDataRow("CO") = vCD.CO
                vDataRow("IE") = vCD.IE
                vDataRow("MainValue") = vCD.MainValue
                vDataRow("SubsidiaryValue") = vCD.SubsidiaryValue
                vDataRow("Period") = vCD.Period
                mvDataSet.Tables("DataRow").Rows.Add(vDataRow)
                .Populate(mvDataSet, False, True)
                SetupCols(mvCurrentSS)
              End If
            End If
          Else
            If mvDataSetDefault IsNot Nothing Then
              If Not mvDataSetDefault.Tables.Contains("DataRow") Then AddDataRowTableToDataSet(mvDataSetDefault)
              If mvDataSetDefault.Tables.Contains("DataRow") Then
                Dim vDataRow As DataRow = mvDataSetDefault.Tables("DataRow").NewRow()
                vDataRow("SearchArea") = vCD.SearchArea
                vDataRow("CO") = vCD.CO
                vDataRow("IE") = vCD.IE
                vDataRow("MainValue") = vCD.MainValue
                vDataRow("SubsidiaryValue") = vCD.SubsidiaryValue
                vDataRow("Period") = vCD.Period
                mvDataSetDefault.Tables("DataRow").Rows.Add(vDataRow)
                .Populate(mvDataSetDefault, False, True)
                SetupCols(mvCurrentSS)
              End If
            End If
          End If
        Else
          Dim vList As New ParameterList
          vList("SearchArea") = vCD.SearchArea
          vList("CO") = vCD.CO
          vList("IE") = vCD.IE
          vList("MainValue") = vCD.MainValue
          vList("SubsidiaryValue") = vCD.SubsidiaryValue
          vList("Period") = vCD.Period
          .UpdateDataRow(vRow, vList)
        End If
        .Visible = True
        If mvCurrentSS Is dgrCriteria Then mvCriteriaGridVisible = True
        If mvCurrentSS Is dgrDefault Then mvDefaultGridVisible = True
      End With
    End If
    SetCriteriaDesc(True)
  End Sub

  Protected Overridable Sub frmEditCriteria_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
    SetPanelLayout()
    If mvSelectionManager Then
      If dgrCriteria.Visible Then dgrCriteria.SetColumnVisible("Counted", False)
      If dgrDefault.Visible Then dgrDefault.SetColumnVisible("Counted", False)
    End If
  End Sub

  Private Sub SetPanelLayout()
    If mvCriteriaGridVisible = False AndAlso mvDefaultGridVisible = True Then
      SplitContainerForGrid.Panel1Collapsed = True
    ElseIf (mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyStandingOrderCancellation OrElse mvMailingInfo.MailingType = CareNetServices.MailingTypes.mtyGAYECancellation) Then
      SplitContainerForGrid.Panel2Collapsed = True
    ElseIf mvCriteriaGridVisible = True AndAlso mvDefaultGridVisible = False Then
      SplitContainerForGrid.Panel2Collapsed = True
    End If
  End Sub

  Private Sub SetOK()
    'Check the first row or all rows of the criteria grid
    'to see if they are blank. Set the OK and count buttons accordingly
    Dim vCriteria As Boolean
    vCriteria = GridRowExists(dgrCriteria) AndAlso dgrCriteria.GetValue(0, "IE").Length > 0

    If mvMailingInfo.Revision > 0 Then
      'If there is a list then enable ok
      cmdOK.Enabled = True
    Else
      cmdOK.Enabled = vCriteria Or Not mvSelectionManager
    End If
    If mvSelectionManager Then cmdCount.Enabled = vCriteria
  End Sub

  Private Function ValidateBrackets(ByVal pDgr As DisplayGrid) As Boolean
    Dim vContinue As Boolean
    Dim vCount As Integer
    Dim vError As Boolean

    vContinue = True
    With pDgr
      If .RowCount > 0 Then
        'does the final row have left brackets but no right brackets?
        Dim vRowIndex As Integer = .RowCount - 1
        If .GetValue(vRowIndex, "LeftParenthesis").Length > 0 Then
          If .GetValue(vRowIndex, "RightParenthesis").Length = 0 Then .SetValue(vRowIndex, "LeftParenthesis", "")
        End If
        'does the first row have right brackets but no left brackets?
        vRowIndex = 0
        If .GetValue(vRowIndex, "LeftParenthesis").Length = 0 Then .SetValue(vRowIndex, "RightParenthesis", "")

        While vRowIndex <= .RowCount - 1 AndAlso vContinue
          If .GetValue(vRowIndex, "IE").Length > 0 Then
            vCount = vCount + .GetValue(vRowIndex, "LeftParenthesis").Length
            vCount = vCount - .GetValue(vRowIndex, "RightParenthesis").Length
          Else
            vContinue = False
          End If
          vRowIndex = vRowIndex + 1
        End While
      End If
    End With
    If vCount <> 0 Then
      vError = True
      ShowInformationMessage(InformationMessages.ImNoMatchParenthesis)   ' The number of left parenthesis does not match the number of right parenthesis
    End If

    Return Not vError
  End Function

  Private Sub RemoveRows()
    With dgrCriteria
      For vRowIndex As Integer = 0 To .RowCount - 1
        If .GetValue(vRowIndex, "IE").Length = 0 AndAlso (.GetValue(vRowIndex, "AndOr").Length > 0 OrElse .GetValue(vRowIndex, "LeftParenthesis").Length > 0 OrElse .GetValue(vRowIndex, "RightParenthesis").Length > 0) Then
          .DeleteRow(vRowIndex)
          If mvDataSet IsNot Nothing Then mvDataSet.AcceptChanges()
        End If
      Next
    End With
  End Sub

  Private Sub dgrDefault_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrDefault.Enter
    Try
      If dgrCriteria.CurrentDataRow > -1 Then dgrCriteria.SelectRow(-1)
      mvCurrentSS = dgrDefault
      ChangeCurrentRow(mvCurrentSS, mvCurrentSS.CurrentDataRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrCriteria_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgrCriteria.Enter
    Try
      If dgrDefault.CurrentDataRow > -1 Then dgrDefault.SelectRow(-1)
      mvCurrentSS = dgrCriteria
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class