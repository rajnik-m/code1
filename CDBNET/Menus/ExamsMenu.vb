Public Class ExamsMenu
  Inherits ExamBaseMenu

  Private mvParent As MaintenanceParentForm
  Private mvExamID As Integer
  Private mvParentID As Integer
  Private mvSessionID As Integer
  Private mvSelector As Boolean
  Private mvItemID As Integer

  Public Event MenuSelected(ByVal sender As Object, ByVal pItem As ExamMenuItems)

  Public Enum ExamMenuItems
    Search
    Reports
    CreateProgramme
    CopyLink
    PasteAsChild
    Share
    RemoveLink
    BookerMailing
    CandidateMailing
    AddScheduleMultiple
    SupplementaryInformation
    Unallocate
    Reallocate
    Customise
    Revert
    Clone
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent

    With mvMenuItems
      .Add(ExamMenuItems.Search.ToString, New MenuToolbarCommand("Search", "Search...", ExamMenuItems.Search, "SCXMUS"))
      .Add(ExamMenuItems.Reports.ToString, New MenuToolbarCommand("Reports", "Reports", ExamMenuItems.Reports, ""))
      .Add(ExamMenuItems.CreateProgramme.ToString, New MenuToolbarCommand("CreateProgramme", "Create Programme", ExamMenuItems.CreateProgramme, "SCXMUC"))
      .Add(ExamMenuItems.CopyLink.ToString, New MenuToolbarCommand("CopyLink", "Copy Link", ExamMenuItems.CopyLink, ""))
      .Add(ExamMenuItems.PasteAsChild.ToString, New MenuToolbarCommand("PasteAsChild", "Paste Link As Child", ExamMenuItems.PasteAsChild, ""))
      .Add(ExamMenuItems.Share.ToString, New MenuToolbarCommand("Share", "Share/Un-share...", ExamMenuItems.Share, "SCXMSU"))
      .Add(ExamMenuItems.RemoveLink.ToString, New MenuToolbarCommand("RemoveLink", "Remove from Parent", ExamMenuItems.RemoveLink, ""))
      .Add(ExamMenuItems.BookerMailing.ToString, New MenuToolbarCommand("BookerMailing", "Booker Mailing", ExamMenuItems.BookerMailing, "SCMMXB"))
      .Add(ExamMenuItems.CandidateMailing.ToString, New MenuToolbarCommand("CandidateMailing", "Candidate Mailing", ExamMenuItems.CandidateMailing, "SCMMXC"))
      .Add(ExamMenuItems.AddScheduleMultiple.ToString, New MenuToolbarCommand("AddSchedule", "Schedules for Multiple Centres", ExamMenuItems.AddScheduleMultiple, ""))
      .Add(ExamMenuItems.SupplementaryInformation.ToString, New MenuToolbarCommand("SupplementaryInformation", "Supplementary Information...", ExamMenuItems.SupplementaryInformation, ""))
      .Add(ExamMenuItems.Unallocate.ToString, New MenuToolbarCommand("Unallocate", "Unallocate", ExamMenuItems.Unallocate, "SCXMMU"))
      .Add(ExamMenuItems.Reallocate.ToString, New MenuToolbarCommand("Reallocate", "Reallocate", ExamMenuItems.Reallocate, "SCXMMR"))
      .Add(ExamMenuItems.Customise.ToString, New MenuToolbarCommand(ExamMenuItems.Customise.ToString, ControlText.MnuDisplayListCustomise, ExamMenuItems.Customise, ""))
      .Add(ExamMenuItems.Revert.ToString, New MenuToolbarCommand(ExamMenuItems.Revert.ToString, ControlText.MnuDisplayListRevert, ExamMenuItems.Revert, ""))
      .Add(ExamMenuItems.Clone.ToString, New MenuToolbarCommand("Clone", "Clone", ExamMenuItems.Clone, "SCXMSD"))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
  End Sub

  Public Sub SetContext(ByVal pDataType As ExamsAccess.XMLExamDataSelectionTypes, ByVal pID As Integer, ByVal pParentID As Integer, ByVal pSessionID As Integer, Optional ByVal pSelector As Boolean = False, Optional ByVal pItemID As Integer = 0)
    ExamDataType = pDataType
    mvExamID = pID
    mvParentID = pParentID
    mvSessionID = pSessionID
    mvSelector = pSelector
    mvItemID = pItemID
  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ExamMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, ExamMenuItems)

      Select Case vMenuItem
        Case ExamMenuItems.CopyLink,
             ExamMenuItems.PasteAsChild,
             ExamMenuItems.RemoveLink,
             ExamMenuItems.Search,
             ExamMenuItems.CreateProgramme,
             ExamMenuItems.AddScheduleMultiple,
             ExamMenuItems.SupplementaryInformation,
             ExamMenuItems.Unallocate,
             ExamMenuItems.Reallocate,
             ExamMenuItems.Share,
             ExamMenuItems.Clone
          RaiseEvent MenuSelected(Me, vMenuItem)
        Case ExamMenuItems.BookerMailing
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyExamBookings, CareServices.TaskJobTypes.tjtExamBookerMailing)
          vGenMail.Process(0)
        Case ExamMenuItems.CandidateMailing
          Dim vGenMail As New GeneralMailing(CareNetServices.MailingTypes.mtyExamCandidates, CareServices.TaskJobTypes.tjtExamCandidateMailing)
          vGenMail.Process(0)
        Case ExamMenuItems.Customise
          MyBase.DoCustomise(False)
        Case ExamMenuItems.Revert
          MyBase.DoCustomise(True)
      End Select
    Catch vException As CareException
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ExamsMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    For Each vItem As ToolStripItem In Me.Items
      vItem.Visible = False
    Next
    Dim vShow As Boolean = False
    Try
      e.Cancel = False
      Me.Items(ExamMenuItems.Unallocate).Visible = False
      Me.Items(ExamMenuItems.Reallocate).Visible = False
      DirectCast(Me.Items(ExamMenuItems.BookerMailing), ToolStripMenuItem).DropDownItems.Clear()
      DirectCast(Me.Items(ExamMenuItems.CandidateMailing), ToolStripMenuItem).DropDownItems.Clear()
      Select Case ExamDataType
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnits
          If mvExamID > 0 Then
            mvMenuItems(ExamMenuItems.Search).SetContextItemVisible(Me, True)
            If mvMenuItems(ExamMenuItems.Search).HideItem = False Then vShow = True
            'No longer required  Me.Items(ExamMenuItems.CopyLink).Visible = True
            If Clipboard.ContainsData(GetType(ExamCopyInfo).FullName) Then
              Dim vExamInfo As ExamCopyInfo = DirectCast(Clipboard.GetData(GetType(ExamCopyInfo).FullName), ExamCopyInfo)
              If vExamInfo IsNot Nothing AndAlso vExamInfo.ExamUnitId <> mvExamID Then
                Me.Items(ExamMenuItems.PasteAsChild).Visible = True
                vShow = True
              End If
            End If
            If mvParentID > 0 Then
              'No longer required Me.Items(ExamMenuItems.RemoveLink).Visible = True
              'We are disabling sharing of top level items.  It doesn't make sense to share a top level item with another item, as it will make the shared item a child in some instances and a top level item in others.
              If ExamsDataHelper.GradingMethod = ExamGradingMethod.NG Then
                mvMenuItems(ExamMenuItems.Share).SetContextItemVisible(Me, (mvSessionID.Equals(0))) 'BR18035 - We will only allowed shared Units for the NG Grading Method.  The Concept Grading Method isn't written to handle this.
                If mvMenuItems(ExamMenuItems.Share).HideItem = False Then vShow = True
              End If
            Else
              If mvSessionID = 0 Then
                mvMenuItems(ExamMenuItems.CreateProgramme).SetContextItemVisible(Me, True)
                If mvMenuItems(ExamMenuItems.CreateProgramme).HideItem = False Then vShow = True
              End If
            End If
            mvMenuItems(ExamMenuItems.BookerMailing).SetContextItemVisible(Me, True)
            mvMenuItems(ExamMenuItems.CandidateMailing).SetContextItemVisible(Me, True)
            If mvMenuItems(ExamMenuItems.BookerMailing).HideItem = False Then vShow = True
            If mvMenuItems(ExamMenuItems.CandidateMailing).HideItem = False Then vShow = True
          End If
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamPersonnel,
             ExamsAccess.XMLExamDataSelectionTypes.ExamCentres,
             ExamsAccess.XMLExamDataSelectionTypes.ExamExemptions
          mvMenuItems(ExamMenuItems.Search).SetContextItemVisible(Me, True)
          If mvMenuItems(ExamMenuItems.Search).HideItem = False Then vShow = True
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamSessions
          mvMenuItems(ExamMenuItems.Search).SetContextItemVisible(Me, True)
          If Not mvMenuItems(ExamMenuItems.Search).HideItem Then
            vShow = True
          End If
          If TypeOf DirectCast(sender, ExamsMenu).SourceControl Is VistaTreeView AndAlso
             DirectCast(DirectCast(sender, ExamsMenu).SourceControl, VistaTreeView).SelectedNode IsNot Nothing Then
            mvMenuItems(ExamMenuItems.Clone).SetContextItemVisible(Me, True)
            If Not mvMenuItems(ExamMenuItems.Clone).HideItem Then
              vShow = True
            End If
          End If
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamSessionCentres,
             ExamsAccess.XMLExamDataSelectionTypes.ExamCentreUnits,
             ExamsAccess.XMLExamDataSelectionTypes.ExamExemptionUnits
          If mvSelector Then
            mvMenuItems(ExamMenuItems.Search).SetContextItemVisible(Me, True)
            If mvMenuItems(ExamMenuItems.Search).HideItem = False Then vShow = True
          End If
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamSchedule
          Me.Items(ExamMenuItems.AddScheduleMultiple).Visible = True
          vShow = True
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCandidateActivites
          Me.Items(ExamMenuItems.SupplementaryInformation).Visible = True
          vShow = True
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamPersonnelMarkerInfo
          Dim vVisible As Boolean = (TypeOf Me.SourceControl Is CDBNETCL.DisplayGrid)
          If vVisible Then
            'Only display if there is data
            Dim vDGR As DisplayGrid = CType(Me.SourceControl, DisplayGrid)
            If vDGR.CurrentRow >= 0 AndAlso vDGR.GetValue(vDGR.CurrentRow, 0).Length = 0 Then vVisible = False
          End If
          mvMenuItems(ExamMenuItems.Reallocate).SetContextItemVisible(Me, vVisible)
          mvMenuItems(ExamMenuItems.Unallocate).SetContextItemVisible(Me, vVisible)
          If vVisible Then
            If mvMenuItems(ExamMenuItems.Reallocate).HideItem = False Then vShow = True
            If mvMenuItems(ExamMenuItems.Unallocate).HideItem = False Then vShow = True
          End If
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamMaintenanceButtons To ExamsAccess.XMLExamDataSelectionTypes.ExamMaintenanceExemptions
          If Customise AndAlso mvExamID = 0 AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDisplayListMaintenance) Then
            mvMenuItems(ExamMenuItems.Customise).SetContextItemVisible(Me, True)
            mvMenuItems(ExamMenuItems.Revert).SetContextItemVisible(Me, True)
            If mvMenuItems(ExamMenuItems.Customise).HideItem = False Then vShow = True
            If mvMenuItems(ExamMenuItems.Revert).HideItem = False Then vShow = True
          End If
      End Select

      DirectCast(Me.Items(ExamMenuItems.Reports), ToolStripMenuItem).DropDownItems.Clear()
      Dim vReportCount As Integer = 0
      Dim vList As New ParameterList(True)
      Dim vCode As String = ""
      Select Case ExamDataType
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamUnits
          vCode = "EXU*"
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamCentres
          vCode = "EXC*"
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamExemptions
          vCode = "EXE*"
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamPersonnel
          vCode = "EXP*"
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamSessions
          vCode = "EXS*"
      End Select
      If vCode.Length > 0 Then
        vList("ReportCode") = vCode
        Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtReports, vList)
        If vTable IsNot Nothing Then
          For Each vRow As DataRow In vTable.Rows
            DirectCast(Me.Items(ExamMenuItems.Reports), ToolStripMenuItem).DropDownItems.Add(vRow("ReportName").ToString, Nothing, AddressOf ReportMenuHandler).Tag = vRow("ReportNumber")
            vReportCount += 1
          Next
        End If
      End If
      mvMenuItems(ExamMenuItems.Reports).SetContextItemVisible(Me, (vReportCount > 0))
      If vReportCount > 0 AndAlso mvMenuItems(ExamMenuItems.Reports).HideItem = False Then vShow = True

      If Not vShow Then e.Cancel = True
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ReportMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Dim vReportNumber As Integer = CInt(vMenuItem.Tag)
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ReportNumber") = vReportNumber
      Select Case ExamDataType
        Case ExamsAccess.XMLExamDataSelectionTypes.ExamPersonnel,
          ExamsAccess.XMLExamDataSelectionTypes.ExamCentres,
          ExamsAccess.XMLExamDataSelectionTypes.ExamSessions,
          ExamsAccess.XMLExamDataSelectionTypes.ExamUnits,
          ExamsAccess.XMLExamDataSelectionTypes.ExamExemptions
          vList.IntegerValue("RP1") = mvItemID
      End Select
      Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class

<Serializable()>
Public Class ExamCopyInfo
  Private mvExamUnitId As Integer

  Sub New(ByVal pExamUnitID As Integer)
    mvExamUnitId = pExamUnitID
  End Sub

  Public ReadOnly Property ExamUnitId As Integer
    Get
      Return mvExamUnitId
    End Get
  End Property
End Class

