Public Class ExamUnitSelector

  Private mvDataType As CareServices.XMLContactDataSelectionTypes
  Private mvContactInfo As ContactInfo
  Private mvItemDataSet As DataSet
  Private mvActivitiesDataSet As DataSet
  Friend WithEvents mvExamsMenu As ExamsMenu
  Private mvExamBookingId As Integer
  Private mvExamStudentHeaderId As Integer
  Private mvExamBookingUnitId As Integer

  Public Property ExamStudentHeaderID As Integer
    Get
      Return mvExamStudentHeaderId
    End Get
    Private Set(value As Integer)
      mvExamStudentHeaderId = value
      If value > 0 Then
        ExamBookingId = 0
        ExamBookingUnitId = 0
      End If
    End Set
  End Property

  Public Event CustomiseCardSet(ByVal Sender As Object, ByVal pDataSelectionType As Integer, ByVal pRevert As Boolean)

  Public Sub New()
    ' This call is required by the designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    spl.FixedPanel = FixedPanel.Panel1
    SetControlTheme()
  End Sub

  Public Sub SetContext(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactInfo As ContactInfo)
    mvDataType = pType
    mvContactInfo = pContactInfo
    sel.Clear()
    dpl.Clear()
    dgr.Clear()
    dgrActivities.Clear()
    dgrGradeHistory.Clear()
    dgrWorkstreams.Clear()
    If dpl.ContextMenuStrip IsNot Nothing Then DirectCast(dpl.ContextMenuStrip, CustomiseMenu).EditMenuVisible = False
    mvItemDataSet = Nothing
  End Sub

  Public Sub InitForSummary(pExamStudentHeaderID As Integer)
    ExamStudentHeaderID = pExamStudentHeaderID
    sel.Init(ExamSelector.SelectionType.StudentCourses, ExamStudentHeaderID)
    sel.SelectNode(0, True)
  End Sub

  Public Sub InitForBooking(pSessionId As Integer, pContactNumber As Integer, pExamBookingId As Integer, ByVal pIsContactExamDetails As Boolean)
    ExamBookingId = pExamBookingId
    sel.Init(ExamSelector.SelectionType.BookingUnits, pSessionId, pContactNumber, ExamBookingId, pIsContactExamDetails:=pIsContactExamDetails)
    sel.SelectNode(0, True)
  End Sub

  Public Sub InitForBooking(pSessionId As Integer, pContactNumber As Integer, pExamBookingId As Integer)
    InitForBooking(pSessionId, pContactNumber, pExamBookingId, False)
  End Sub

  Private Sub sel_ItemSelected(ByVal sender As Object, ByVal pType As CDBNETCL.ExamsAccess.XMLExamDataSelectionTypes, ByVal pItem As ExamSelectorItem) Handles sel.ItemSelected
    Try
      Dim vList As New ParameterList(True)
      Select Case mvDataType
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummary
          'For Student Courses pItemId is UnitId GetID2 is ExamStudentUnitHeaderId
          Dim vExamStudentUnitHeaderId As Integer = DirectCast(sender, ExamSelector).GetID2
          vList.IntegerValue("ExamStudentUnitHeaderId") = DirectCast(sender, ExamSelector).GetID2
          mvItemDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactExamSummaryItems, mvContactInfo.ContactNumber, vList)
          dpl.Init(mvItemDataSet, False, False)
          dpl.Populate(mvItemDataSet, 0)
          DirectCast(dpl.ContextMenuStrip, CustomiseMenu).DataSelectionType = dpl.DataSelectionType
          DirectCast(dpl.ContextMenuStrip, CustomiseMenu).CancelMenuVisible = False
          DirectCast(dpl.ContextMenuStrip, CustomiseMenu).EditMenuVisible = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciContactEditExamResults)
          vList.IntegerValue("ExamUnitId") = DirectCast(sender, ExamSelector).GetUnitID
          vList.Remove("ExamStudentUnitHeaderId")
          Dim vDataSet As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamSummaryList, mvContactInfo.ContactNumber, vList)
          dgr.Populate(vDataSet)
          'Grade History
          If Not tab.Controls.Contains(tbp4) Then tab.Controls.Add(tbp4)
          If vList.ContainsKey("ExamUnitId") AndAlso pItem IsNot Nothing Then vList.IntegerValue("ExamUnitId") = pItem.UnitID
          If vExamStudentUnitHeaderId > 0 AndAlso vList.Contains("ExamStudentUnitHeaderId") = False Then vList.IntegerValue("ExamStudentUnitHeaderId") = vExamStudentUnitHeaderId
          vDataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamUnitGradeHistory, vList)
          dgrGradeHistory.Populate(vDataSet)
          Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
          If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            'Hide grid
            tab.Controls.Remove(tbp4)
          End If
          If tab.Controls.Contains(tbp3) Then tab.Controls.Remove(tbp3)
          If tab.Controls.Contains(tbp5) Then tab.Controls.Remove(tbp5)
        Case CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetails
          'For booking units pItemId is UnitId GetID2 is ExamBookingId
          ExamBookingUnitId = DirectCast(sender, ExamSelector).GetID2
          vList.IntegerValue("ExamBookingUnitId") = ExamBookingUnitId
          mvItemDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactExamDetailItems, mvContactInfo.ContactNumber, vList)
          Dim vOutOfContextBookingUnit = Not IsCurrentUnitInBooking()
          If vOutOfContextBookingUnit Then
            AddWarningLabelControl()
          End If
          dpl.ColumnSizingType = DisplayPanel.ColumnSizingTypes.Automatic
          dpl.Init(mvItemDataSet, False, False)
          dpl.Populate(mvItemDataSet, 0)
          DirectCast(dpl.ContextMenuStrip, CustomiseMenu).DataSelectionType = dpl.DataSelectionType
          DirectCast(dpl.ContextMenuStrip, CustomiseMenu).EditMenuVisible = AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciContactEditExamResults)
          Dim vCanCancel As Boolean = DataHelper.UserInfo.AccessLevel = UserInfo.UserAccessLevel.ualDatabaseAdministrator
          Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(mvItemDataSet)
          If vDataRow.Item("CancellationReason").ToString.Length > 0 Then vCanCancel = False
          'mvExamBookingId = IntegerValue(vDataRow("ExamBookingId").ToString)
          'Blank out Booking controls if the selected bookings unit doesn't belong to the Booking (e.g. split booking)
          If vOutOfContextBookingUnit Then
            ObscureNonContextBooking()
            ShowWarningLabel(IntegerValue(vDataRow("ExamBookingId").ToString))
          End If
          DirectCast(dpl.ContextMenuStrip, CustomiseMenu).CancelMenuVisible = vCanCancel
          vList.Remove("ExamBookingUnitId")
          vList.IntegerValue("ExamUnitId") = If(pItem IsNot Nothing, pItem.UnitID, 0)
          vList.IntegerValue("ExamUnitLinkId") = If(pItem IsNot Nothing, pItem.LinkID, 0)
          vList.IntegerValue("ExamBookingId") = ExamBookingId
          Dim vDataSet As DataSet = DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExamDetailList, mvContactInfo.ContactNumber, vList)
          dgr.Populate(vDataSet)


          'Categories
          mvActivitiesDataSet = Nothing
          mvExamsMenu = Nothing
          If Not vOutOfContextBookingUnit Then
            mvActivitiesDataSet = Nothing
            mvExamsMenu = Nothing
            Dim vActivityParams As New ParameterList(True)
            vActivityParams.AddSystemColumns()
            vActivityParams.Add("ExamBookingUnitId", ExamBookingUnitId)
            mvActivitiesDataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamCandidateActivites, vActivityParams)
            dgrActivities.AutoSetHeight = True
            dgrActivities.Populate(mvActivitiesDataSet)
            mvExamsMenu = New ExamsMenu(CType(Me.ParentForm, MaintenanceParentForm))
            dgrActivities.ContextMenuStrip = mvExamsMenu
            mvExamsMenu.SetContext(ExamsAccess.XMLExamDataSelectionTypes.ExamCandidateActivites, DirectCast(sender, ExamSelector).GetID2, 0, 0)
            dgrActivities.SetToolBarVisible()
          End If

          'Grade History
          Dim vAddHistoryTab As Boolean = False 'only add the history if it has data
          If Not vOutOfContextBookingUnit Then
            If vList.ContainsKey("ExamUnitId") AndAlso pItem IsNot Nothing Then vList.IntegerValue("ExamUnitId") = pItem.UnitID
            If ExamBookingUnitId > 0 AndAlso vList.Contains("ExamBookingUnitId") = False Then vList.IntegerValue("ExamBookingUnitId") = ExamBookingUnitId
            vDataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamUnitGradeHistory, vList)
            dgrGradeHistory.Populate(vDataSet)
            Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
            If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
              vAddHistoryTab = True
            End If
          End If

          'Workstreams
          Dim vAddWorkstreamsTab As Boolean = False
          If Not vOutOfContextBookingUnit Then
            If ExamBookingUnitId > 0 Then
              Dim vWorkstreamParameters As New ParameterList(True, True)
              vWorkstreamParameters.IntegerValue("ExamBookingUnitId") = ExamBookingUnitId
              vDataSet = WorkstreamDataHelper.SelectWorkstreamDataSet(WorkstreamService.XMLDataSelectionTypes.WorkstreamDetails, vWorkstreamParameters)
              dgrWorkstreams.Populate(vDataSet)
              Dim vDT As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
              If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
                dgrWorkstreams.SetToolBarVisible()
                vAddWorkstreamsTab = True
              End If
            End If
          End If

          If vOutOfContextBookingUnit Then
            If tab.Controls.Contains(tbp3) Then tab.Controls.Remove(tbp3)
            If tab.Controls.Contains(tbp4) Then tab.Controls.Remove(tbp4)
          Else
            If Not tab.Controls.Contains(tbp3) Then tab.Controls.Add(tbp3)
            If vAddHistoryTab Then
              If Not tab.Controls.Contains(tbp4) Then tab.Controls.Add(tbp4)
            Else
              If tab.Controls.Contains(tbp4) Then tab.Controls.Remove(tbp4)
            End If
            If vAddWorkstreamsTab Then
              If Not tab.Controls.Contains(tbp5) Then tab.Controls.Add(tbp5)
            Else
              If tab.Controls.Contains(tbp5) Then tab.Controls.Remove(tbp5)
            End If
          End If

      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Public Property DisplayPanelContextMenu As ContextMenuStrip
    Get
      Return dpl.ContextMenuStrip
    End Get
    Set(ByVal value As ContextMenuStrip)
      dpl.ContextMenuStrip = value
    End Set
  End Property

  Public Property ExamBookingId As Integer
    Get
      Return mvExamBookingId
    End Get
    Private Set(value As Integer)
      mvExamBookingId = value
      If value > 0 Then ExamStudentHeaderID = 0
    End Set
  End Property

  Public Property ExamBookingUnitId As Integer
    Get
      Return mvExamBookingUnitId
    End Get
    Private Set(value As Integer)
      mvExamBookingUnitId = value
      If value > 0 Then ExamStudentHeaderID = 0
    End Set
  End Property

  Public ReadOnly Property ExamSelector As ExamSelector
    Get
      Return sel
    End Get
  End Property

  Public ReadOnly Property ItemDataSet As DataSet
    Get
      Return mvItemDataSet
    End Get
  End Property

  Private Sub dgr_CustomiseCardSet(ByVal Sender As Object, ByVal pDataSelectionType As Integer, ByVal pRevert As Boolean) Handles dgr.CustomiseCardSet, dgrGradeHistory.CustomiseCardSet, dgrWorkstreams.CustomiseCardSet, dgrActivities.CustomiseCardSet
    RaiseEvent CustomiseCardSet(Sender, pDataSelectionType, pRevert)
  End Sub

  Private Sub mvExamsMenu_MenuSelected(ByVal sender As Object, ByVal pItem As ExamsMenu.ExamMenuItems) Handles mvExamsMenu.MenuSelected
    Select Case pItem
      Case ExamsMenu.ExamMenuItems.SupplementaryInformation
        If sel.GetID2 > 0 Then
          Dim vFoundRows() As Data.DataRow
          Dim vRow As Data.DataRow

          If mvItemDataSet IsNot Nothing Then
            ' Find the exam booking unit in the dataset for the activity group and source
            vFoundRows = DataHelper.GetTableFromDataSet(mvItemDataSet).Select("ExamBookingUnitId = " + sel.GetID2.ToString)
            If vFoundRows.Length = 1 Then
              vRow = vFoundRows(0)
              ShowExamCandidateDataSheet(Nothing, mvContactInfo, "C", vRow.Item("Source").ToString(), vRow.Item("ActivityGroup").ToString(), vRow.Item("ExamUnitDescription").ToString(), sel.GetID2, False)
              sel.ReSelectNode()
            End If
          End If
        End If
    End Select
  End Sub
  ''' <summary>
  ''' Blanks out all booking-related controls.  Should be called if the selected Unit is not part of the selected Booking.
  ''' </summary>
  ''' <remarks>
  ''' This is only relevant when split bookings are present:  when a Booking is made for an Exam Unit, all parent units for the booking are created as Exam Booking Units.  If a subsequent booking is then made for additional units (before Grading is run) i.e. the units are split over two bookings, the Parent Units are not created again as that would duplicate the parents
  ''' over the two bookings.  Instead, the tree displays the parent regardless of what original booking was, and then displays only the children that applies to the booking.
  ''' This is problematic when viewing data as the user might expect the parent units to be duplicated, and for example have two possible different grades.
  ''' To make this more visible, we hide the booking-related units when the 
  ''' </remarks>
  Private Sub ObscureNonContextBooking()

    'The current unit doesn't belong to the Booking that the user has selected (see remarks above)
    Dim vHiddenControls() As String = {"ExamCandidateNumber", "AttemptNumber", "ExamStudentUnitStatus", "ExamStudentUnitStatusDesc", _
                                       "OriginalMark", "ModeratedMark", "TotalMark", "OriginalGrade", "ModeratedGrade", "TotalGrade", _
                                       "TotalGradeDesc", "OriginalResult", "ModeratedResult", "TotalResult", "TotalResultDesc", "DoneDate", _
                                       "StartDate", "StartTime", "EndTime", "Source", "ActivityGroup", "CancellationReason", "CancellationSource", _
                                       "CancellationReasonDesc", "CancellationSourceDesc", "CancelledOn", "CancelledBy", "ExamCentreId", "CanEditResults", _
                                       "CourseStartDate", "ExamAssessmentLanguage", "StudyMode", "CentreUnitLocalName"}
    Dim vMessage As String = "n/a"
    For Each vItem As String In vHiddenControls
      Dim vControl As Control = FindControl(dpl, vItem, False)
      If vControl IsNot Nothing Then
        vControl.Text = vMessage
        vControl.Enabled = False
      End If
    Next

  End Sub

  Private Sub AddWarningLabelControl()
    Dim vDT As DataTable = DataHelper.GetColumnTableFromDataSet(mvItemDataSet)
    If vDT IsNot Nothing AndAlso vDT.Rows IsNot Nothing Then
      If vDT.Rows.Cast(Of DataRow).FirstOrDefault(Function(vRow As DataRow) vRow("Name").ToString = "InformationMessage") Is Nothing Then
        Dim vRow As DataRow = vDT.NewRow()
        'vRow("Name") = "Spacer_InformationMessage"
        'vRow("Value") = ""
        'vRow("Visible") = False
        'vRow("DataType") = "Char"
        'vRow("Heading") = ""
        'vDT.Rows.Add(vRow)

        'vRow = vDT.NewRow()
        vRow("Name") = "Warning_LabelOnly"
        vRow("Value") = ""
        vRow("Visible") = False
        vRow("DataType") = "Char"
        vRow("Heading") = "Warning:"
        vDT.Rows.Add(vRow)

        vRow = vDT.NewRow()
        vRow("Name") = "InformationMessage_LabelOnly"
        vRow("Value") = ""
        vRow("Visible") = False
        vRow("DataType") = "Char"
        vRow("Heading") = "  "
        vDT.Rows.Add(vRow)

        vRow = DataHelper.GetRowFromDataSet(mvItemDataSet)

        If vRow IsNot Nothing Then
          vRow.Table.Columns.Add("Warning_LabelOnly", GetType(String))
          vRow.Table.Columns.Add("InformationMessage_LabelOnly", GetType(String))
        End If

      End If
    End If
  End Sub
  Private Sub ShowWarningLabel(pOutOfContextBooking As Integer)
    Dim vInfoLabel As Label = dpl.FindLabel("InformationMessage_LabelOnly")
    Dim vWarningLabel As Label = dpl.FindLabel("Warning_LabelOnly")
    If vInfoLabel IsNot Nothing Then
      vInfoLabel.Visible = False
      vWarningLabel.Visible = False
    End If
    If vWarningLabel IsNot Nothing Then
      vWarningLabel.Font = New Font(vWarningLabel.Font, FontStyle.Bold)
      vWarningLabel.Visible = True
    End If

    If vInfoLabel IsNot Nothing Then
      vInfoLabel.Text = String.Format(InformationMessages.ImExamBookingEditFromOtherBooking, pOutOfContextBooking.ToString)
      vInfoLabel.Visible = True
    End If

  End Sub

  ''' <summary>
  ''' Returns True if the unit that is selected in the tree is part of the booking that was passed in the initialiser
  ''' </summary>
  ''' <returns>True if the booking Id for the Unit is the same at the Booking Id that was given to the Exam Unit Selector, False if not.</returns>
  ''' <remarks>In a tree with multiple children, it is possible to book each child in a separate booking.  When this occurs, the units required to build the tree are booked 
  ''' as part of the first booking.  Subsequent bookings only own the additional units.  This can lead to confusion because when a Booking is selected in the grid, the Exam 
  ''' tree can display parent units that actually belong to a different booking.  To remedy this, the control displays warning messages, hides data and tabs and prevents edits.
  ''' All of this happens when the method below returns False</remarks>
  Private Function IsCurrentUnitInBooking() As Boolean
    Dim vRtn As Boolean = True
    If ItemDataSet IsNot Nothing Then
      Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(ItemDataSet)
      If vDataRow IsNot Nothing Then
        Dim vContextBookingId As Integer = IntegerValue(vDataRow("ExamBookingId").ToString)
        vRtn = (vContextBookingId = ExamBookingId)
      End If
    End If

    Return vRtn
  End Function

  Private Sub SetControlTheme()
    For Each vControl As Control In Me.Controls
      DoSetControlTheme(vControl)
    Next
  End Sub

End Class
