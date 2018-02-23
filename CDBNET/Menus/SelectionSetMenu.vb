Public Class SelectionSetMenu
  Inherits ContextMenuStrip

  Private mvParent As MaintenanceParentForm
  Private mvSelectionSetNumber As Integer

  Private Enum SelectionSetMenuItems
    ssmiNew
    ssmiDelete
    ssmiReport
    ssmiMailing
    ssmiMerge
    ssmiRename
    ssmiCopy
    ssmiDeleteAllContacts
    ssmiSurveyRegistration
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent
    With mvMenuItems
      .Add(SelectionSetMenuItems.ssmiNew.ToString, New MenuToolbarCommand("New", ControlText.MnuSelectionSetNew, SelectionSetMenuItems.ssmiNew))
      .Add(SelectionSetMenuItems.ssmiReport.ToString, New MenuToolbarCommand("Report", ControlText.MnuSelectionSetReport, SelectionSetMenuItems.ssmiReport))
      .Add(SelectionSetMenuItems.ssmiMailing.ToString, New MenuToolbarCommand("Mailing", ControlText.MnuSelectionSetMailing, SelectionSetMenuItems.ssmiMailing))
      .Add(SelectionSetMenuItems.ssmiDelete.ToString, New MenuToolbarCommand("Delete", ControlText.MnuSelectionSetDelete, SelectionSetMenuItems.ssmiDelete))
      .Add(SelectionSetMenuItems.ssmiMerge.ToString, New MenuToolbarCommand("Merge", ControlText.MnuSelectionSetMerge, SelectionSetMenuItems.ssmiMerge))
      .Add(SelectionSetMenuItems.ssmiRename.ToString, New MenuToolbarCommand("Rename", ControlText.MnuSelectionSetRename, SelectionSetMenuItems.ssmiRename))
      .Add(SelectionSetMenuItems.ssmiCopy.ToString, New MenuToolbarCommand("Copy", ControlText.MnuSelectionSetCopy, SelectionSetMenuItems.ssmiCopy))
      .Add(SelectionSetMenuItems.ssmiDeleteAllContacts.ToString, New MenuToolbarCommand("DeleteAllContacts", ControlText.MnuSelectionSetDeleteAllContacts, SelectionSetMenuItems.ssmiDeleteAllContacts, "SCSSDA"))
      .Add(SelectionSetMenuItems.ssmiSurveyRegistration.ToString, New MenuToolbarCommand("SurveyRegistration", ControlText.MnuSurveyRegistration, SelectionSetMenuItems.ssmiSurveyRegistration, "SCSSSR"))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
  End Sub

  Public Property SelectionSetNumber() As Integer
    Get
      Return mvSelectionSetNumber
    End Get
    Set(ByVal Value As Integer)
      mvSelectionSetNumber = Value
    End Set
  End Property

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As SelectionSetMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, SelectionSetMenuItems)

      Select Case vMenuItem
        Case SelectionSetMenuItems.ssmiNew
          If FormHelper.AddSelectionSet(mvParent) > 0 Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
        Case SelectionSetMenuItems.ssmiReport
          Dim vForm As New frmReportDataSelection(mvSelectionSetNumber, False)
          vForm.ShowDialog()
        Case SelectionSetMenuItems.ssmiMailing
          AppHelper.ProcessSelectionSetMailing(MainHelper.MainForm, mvSelectionSetNumber, True)
        Case SelectionSetMenuItems.ssmiDelete
          If Not ConfirmDelete() Then Exit Sub
          Dim vList As ParameterList = New ParameterList(True)
          vList.IntegerValue("SelectionSetNumber") = mvSelectionSetNumber
          DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vList)
          UserHistory.RemoveOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, mvSelectionSetNumber)
          mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
        Case SelectionSetMenuItems.ssmiDeleteAllContacts
          FormHelper.DoBulkContactDeletion(mvSelectionSetNumber)
        Case SelectionSetMenuItems.ssmiCopy
          Dim vList As New ParameterList(True)
          Dim vParamList As New ParameterList(True)
          vParamList("SelectionSet") = mvSelectionSetNumber.ToString
          Dim vSelectionSetTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSelectionSets, vParamList)
          If Not vSelectionSetTable Is Nothing Then
            vList("SelectionSetDesc") = vSelectionSetTable.Rows(0).Item("SelectionSetDesc").ToString
            vParamList.Remove("SelectionSet")
            'Open Application Parameters form
            vParamList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optCopySelectionSet, Nothing, vList, "Copy Selection Set")
            If Not vParamList Is Nothing Then
              vParamList("SelectionSetNumber") = mvSelectionSetNumber.ToString
              'Call WebService to Copy SelectionSet
                Dim vSSData As DataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstSelectionSetData, CareServices.XMLTableDataSelectionTypes), vParamList)
                If vSSData.Tables("DataRow").Rows.Count > 0 Then
                  vList("OldSelectionSetNumber") = mvSelectionSetNumber.ToString
                  vList("NumberInMailing") = vSSData.Tables("DataRow").Rows(0).Item("NumberInSet").ToString
                  vList("SelectionSetDesc") = vParamList("SelectionSetDesc")
                  vList("CopySelectionSet") = "Y"
                End If
                DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vList)
              End If
              mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
            End If
        Case SelectionSetMenuItems.ssmiRename
          Dim vList As New ParameterList(True)
          Dim vParamList As New ParameterList(True)
          vParamList("SelectionSet") = mvSelectionSetNumber.ToString
          Dim vSelectionSetTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSelectionSets, vParamList)
          If Not vSelectionSetTable Is Nothing Then
            vList("SelectionSetDesc") = vSelectionSetTable.Rows(0).Item("SelectionSetDesc").ToString
            vParamList.Remove("SelectionSet")
            'Open Application Parameters form
            vParamList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optRenameSelectionSet, Nothing, vList, "Rename Selection Set")
            If Not vParamList Is Nothing Then
              vParamList("SelectionSetNumber") = mvSelectionSetNumber.ToString
              'Call WebService to Rename SelectionSet
              DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vParamList)
              mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
          End If
          End If
        Case SelectionSetMenuItems.ssmiMerge
          If FormHelper.MergeSelectionSet(mvParent, mvSelectionSetNumber) > 0 Then mvParent.RefreshData(CareServices.XMLMaintenanceControlTypes.xmctMergeSelectionSet)
        Case SelectionSetMenuItems.ssmiSurveyRegistration
          Dim vList As New ParameterList(True)
          vList("SelectionSet") = mvSelectionSetNumber.ToString
          FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtRegisterSurvey, vList, True, FormHelper.ProcessTaskScheduleType.ptsAskToSchedule, True)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub SelectionSetMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    Dim vCursor As New BusyCursor
    Try
      'make DeletAllContacts visible Only if the User has Access Rights
      Me.Items(SelectionSetMenuItems.ssmiDeleteAllContacts).Visible = Not mvMenuItems.Item(SelectionSetMenuItems.ssmiDeleteAllContacts.ToString).HideItem
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class
