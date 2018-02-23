Public Class EventMenu
  Inherits EventBaseMenu

  Private mvParent As MaintenanceParentForm
  '  Private mvEventDataType As CareServices.XMLEventDataSelectionTypes = CType(-1, CareServices.XMLEventDataSelectionTypes)
  Private mvEventInfo As CareEventInfo

  Public Event MenuSelected(ByVal pItem As EventMenuItems)

  Public Enum EventMenuItems
    emiCalculateEventTotals
    emiCalculateEventAndDelegateTotals
    emiCalculateDelegateTotals
    emiProcessWaitingList
    emiBookerMailing
    emiDelegateMailing
    emiPersonnelMailing
    emiSponsorMailing
    emiAllocatePISToEvent
    emiAllocatePISToDelegates
    emiDuplicateEvent
    emiNumberCandidates
    emiNumberSessionBookings
    emiLoanItems
    emiAuthoriseExpenses
    emiIssueResources
    emiConfirmDelegates
    emiAttendeeReport
    emiAttendeeReport2
    emiDelegateSessionReport
    emiCustomise
    emiRevert
    emiBookingsReport
    emiAccommodationReport
    emiOrganiserReport
    emiNone
  End Enum

  'Public Property EventDataType() As CareServices.XMLEventDataSelectionTypes
  '  Get
  '    Return mvEventDataType
  '  End Get
  '  Set(ByVal pValue As CareServices.XMLEventDataSelectionTypes)
  '    mvEventDataType = pValue
  '  End Set
  'End Property

  Public Property EventInfo() As CareEventInfo
    Get
      Return mvEventInfo
    End Get
    Set(ByVal pValue As CareEventInfo)
      mvEventInfo = pValue
    End Set
  End Property


  Protected mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent

    With mvMenuItems
      .Add(EventMenuItems.emiCalculateEventTotals.ToString, New MenuToolbarCommand("CalculateEvent", ControlText.mnuEventCalculate, EventMenuItems.emiCalculateEventTotals, ""))
      .Add(EventMenuItems.emiCalculateEventAndDelegateTotals.ToString, New MenuToolbarCommand("CalculateAll", ControlText.mnuEventCalculateAllTotals, EventMenuItems.emiCalculateEventAndDelegateTotals, ""))
      .Add(EventMenuItems.emiCalculateDelegateTotals.ToString, New MenuToolbarCommand("CalculateDelegates", ControlText.mnuEventDelegateTotals, EventMenuItems.emiCalculateDelegateTotals, ""))
      .Add(EventMenuItems.emiProcessWaitingList.ToString, New MenuToolbarCommand("ProcessWaitingList", ControlText.mnuProcessWaitingList, EventMenuItems.emiProcessWaitingList, ""))
      .Add(EventMenuItems.emiBookerMailing.ToString, New MenuToolbarCommand("BookerMailing", ControlText.mnuEventBookerMailing, EventMenuItems.emiBookerMailing, ""))
      .Add(EventMenuItems.emiDelegateMailing.ToString, New MenuToolbarCommand("DelegateMailing", ControlText.mnuEventDelegateMailing, EventMenuItems.emiDelegateMailing, ""))
      .Add(EventMenuItems.emiPersonnelMailing.ToString, New MenuToolbarCommand("PersonnelMailing", ControlText.mnuEventPersonnelMailing, EventMenuItems.emiPersonnelMailing, ""))
      .Add(EventMenuItems.emiSponsorMailing.ToString, New MenuToolbarCommand("SponsorMailing", ControlText.mnuEventSponsorMailing, EventMenuItems.emiSponsorMailing, ""))
      .Add(EventMenuItems.emiAllocatePISToEvent.ToString, New MenuToolbarCommand("AllocatePISToEvent", ControlText.mnuAllocatePISToEvent, EventMenuItems.emiAllocatePISToEvent, ""))
      .Add(EventMenuItems.emiAllocatePISToDelegates.ToString, New MenuToolbarCommand("AllocatePISToDelegates", ControlText.mnuAllocatePISToDelegates, EventMenuItems.emiAllocatePISToDelegates, ""))
      .Add(EventMenuItems.emiDuplicateEvent.ToString, New MenuToolbarCommand("DuplicateEvent", ControlText.MnuDuplicateEvent, EventMenuItems.emiDuplicateEvent, ""))
      .Add(EventMenuItems.emiNumberCandidates.ToString, New MenuToolbarCommand("NumberCandidates", ControlText.MnuNumberCandidates, EventMenuItems.emiNumberCandidates, ""))
      .Add(EventMenuItems.emiNumberSessionBookings.ToString, New MenuToolbarCommand("NumberSessionBookings", ControlText.MnuNumberSessionBookings, EventMenuItems.emiNumberSessionBookings, ""))
      .Add(EventMenuItems.emiLoanItems.ToString, New MenuToolbarCommand("LoanItems", ControlText.MnuLoanItems, EventMenuItems.emiLoanItems, ""))
      .Add(EventMenuItems.emiAuthoriseExpenses.ToString, New MenuToolbarCommand("AuthoriseExpenses", ControlText.MnuAuthoriseExpenses, EventMenuItems.emiAuthoriseExpenses, ""))
      .Add(EventMenuItems.emiIssueResources.ToString, New MenuToolbarCommand("IssueResources", ControlText.MnuIssueResources, EventMenuItems.emiIssueResources, "EGIR", , ""))
      .Add(EventMenuItems.emiConfirmDelegates.ToString, New MenuToolbarCommand("ConfirmDelegates", ControlText.MnuEventConfirmDelegates, EventMenuItems.emiConfirmDelegates, ""))
      .Add(EventMenuItems.emiAttendeeReport.ToString, New MenuToolbarCommand("AttendeeReport", ControlText.MnuEventAttendeeReport, EventMenuItems.emiAttendeeReport, ""))
      .Add(EventMenuItems.emiAttendeeReport2.ToString, New MenuToolbarCommand("AttendeeReport2", ControlText.MnuEventAttendeeReport2, EventMenuItems.emiAttendeeReport2, ""))
      .Add(EventMenuItems.emiDelegateSessionReport.ToString, New MenuToolbarCommand("DelegateSessions", ControlText.MnuEventDelegateSessionReport, EventMenuItems.emiDelegateSessionReport, ""))
      .Add(EventMenuItems.emiCustomise.ToString, New MenuToolbarCommand("Customise", ControlText.MnuDisplayListCustomise, EventMenuItems.emiCustomise, ""))
      .Add(EventMenuItems.emiRevert.ToString, New MenuToolbarCommand("Revert", ControlText.MnuDisplayListRevert, EventMenuItems.emiRevert, ""))
      
      .Add(EventMenuItems.emiBookingsReport.ToString, New MenuToolbarCommand("BookingsReport", ControlText.MnuEventBookingsReport, EventMenuItems.emiBookingsReport, ""))
      .Add(EventMenuItems.emiAccommodationReport.ToString, New MenuToolbarCommand("AccommodationReport", ControlText.MnuEventAccommodationReport, EventMenuItems.emiAccommodationReport, ""))
      .Add(EventMenuItems.emiOrganiserReport.ToString, New MenuToolbarCommand("OrganiserReport", ControlText.MnuEventOrganiserReport, EventMenuItems.emiOrganiserReport, ""))
    End With

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)

  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As EventMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, EventMenuItems)

      Select Case vMenuItem
        Case EventMenuItems.emiCalculateEventTotals, EventMenuItems.emiCalculateEventAndDelegateTotals, _
             EventMenuItems.emiCalculateDelegateTotals, EventMenuItems.emiProcessWaitingList, _
             EventMenuItems.emiAllocatePISToEvent, EventMenuItems.emiAllocatePISToDelegates, EventMenuItems.emiDuplicateEvent, _
             EventMenuItems.emiNumberCandidates, EventMenuItems.emiNumberSessionBookings, _
             EventMenuItems.emiConfirmDelegates, EventMenuItems.emiLoanItems, EventMenuItems.emiAuthoriseExpenses, EventMenuItems.emiIssueResources,
             EventMenuItems.emiCustomise, EventMenuItems.emiRevert
          RaiseEvent MenuSelected(vMenuItem)
        Case EventMenuItems.emiAttendeeReport
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "EVATTS"
          vList("RP1") = mvEventInfo.EventNumber.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case EventMenuItems.emiAttendeeReport2
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "EVATTF"
          vList("RP1") = mvEventInfo.EventNumber.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case EventMenuItems.emiDelegateSessionReport
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "EVDELS"
          vList("RP1") = mvEventInfo.EventNumber.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case EventMenuItems.emiBookingsReport
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "EVBOOK"
          vList("RP1") = mvEventInfo.EventNumber.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case EventMenuItems.emiAccommodationReport
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "EVACOM"
          vList("RP1") = mvEventInfo.EventNumber.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case EventMenuItems.emiOrganiserReport
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "EVORGN"
          vList("RP1") = mvEventInfo.EventNumber.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Protected Overridable Sub EventMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    For Each vItem As ToolStripItem In Me.Items
      vItem.Visible = False
    Next
    Dim vMailingCount As Integer
    Dim vCanProcessWaiting As Boolean = (mvEventInfo.NumberOnWaitingList > 0) AndAlso (mvEventInfo.NumberOfAttendees < mvEventInfo.MaximumAttendees)
    DirectCast(Me.Items(EventMenuItems.emiBookerMailing), ToolStripMenuItem).DropDownItems.Clear()
    DirectCast(Me.Items(EventMenuItems.emiDelegateMailing), ToolStripMenuItem).DropDownItems.Clear()
    DirectCast(Me.Items(EventMenuItems.emiPersonnelMailing), ToolStripMenuItem).DropDownItems.Clear()
    DirectCast(Me.Items(EventMenuItems.emiSponsorMailing), ToolStripMenuItem).DropDownItems.Clear()
    If Customise AndAlso AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDisplayListMaintenance) Then
      mvMenuItems(EventMenuItems.emiCustomise).SetContextItemVisible(Me, True)
      mvMenuItems(EventMenuItems.emiRevert).SetContextItemVisible(Me, True)
    End If
    
    Try
      e.Cancel = False
      Select Case EventDataType
        Case CareServices.XMLEventDataSelectionTypes.xedtEventInformation
          mvMenuItems(EventMenuItems.emiCalculateEventTotals).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiCalculateEventAndDelegateTotals).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiDuplicateEvent).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiNumberCandidates).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiNumberSessionBookings).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiLoanItems).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiAuthoriseExpenses).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiIssueResources).SetContextItemVisible(Me, True)

          If mvEventInfo.ContainsBookings Then mvMenuItems(EventMenuItems.emiBookingsReport).SetContextItemVisible(Me, True)
          If mvEventInfo.External AndAlso mvEventInfo.Organiser.Length > 0 Then mvMenuItems(EventMenuItems.emiOrganiserReport).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiAccommodationReport).SetContextItemVisible(Me, True)

          If vCanProcessWaiting Then mvMenuItems(EventMenuItems.emiProcessWaitingList).SetContextItemVisible(Me, True)
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
          vMailingCount = GetMailings(EventMenuItems.emiBookerMailing)
          If vMailingCount > 0 Then mvMenuItems(EventMenuItems.emiBookerMailing).SetContextItemVisible(Me, True)
          If vCanProcessWaiting Then
            mvMenuItems(EventMenuItems.emiProcessWaitingList).SetContextItemVisible(Me, True)
          ElseIf vMailingCount = 0 Then
            e.Cancel = True
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventAttendees
          vMailingCount = GetMailings(EventMenuItems.emiDelegateMailing)
          If vMailingCount > 0 Then mvMenuItems(EventMenuItems.emiDelegateMailing).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiCalculateDelegateTotals).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiConfirmDelegates).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiAttendeeReport).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiAttendeeReport2).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiDelegateSessionReport).SetContextItemVisible(Me, True)
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
          vMailingCount = GetMailings(EventMenuItems.emiPersonnelMailing)
          If vMailingCount > 0 Then
            mvMenuItems(EventMenuItems.emiPersonnelMailing).SetContextItemVisible(Me, True)
          Else
            If Not Customise Then e.Cancel = True
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventCosts
          vMailingCount = GetMailings(EventMenuItems.emiSponsorMailing)
          If vMailingCount > 0 Then
            mvMenuItems(EventMenuItems.emiSponsorMailing).SetContextItemVisible(Me, True)
          Else
            If Not Customise Then e.Cancel = True
          End If
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPIS
          mvMenuItems(EventMenuItems.emiAllocatePISToEvent).SetContextItemVisible(Me, True)
          mvMenuItems(EventMenuItems.emiAllocatePISToDelegates).SetContextItemVisible(Me, True)
        Case Else
          If Not Customise Then e.Cancel = True
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Function GetMailings(ByVal pMenuItem As EventMenuItems) As Integer
    Dim vList As New ParameterList(True)
    Dim vCode As String = ""
    Select Case pMenuItem
      Case EventMenuItems.emiBookerMailing
        vCode = "EB"
      Case EventMenuItems.emiDelegateMailing
        vCode = "EA"
      Case EventMenuItems.emiPersonnelMailing
        vCode = "EP"
      Case EventMenuItems.emiSponsorMailing
        vCode = "ES"
    End Select
    vList("ApplicationName") = vCode
    vList("FieldName") = "event_number"
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCriteriaSets, vList)
    Dim vReports As ToolStripMenuItem = mvMenuItems(pMenuItem).FindMenuStripItem(Me)

    If vTable IsNot Nothing And vReports IsNot Nothing Then
      For Each vRow As DataRow In vTable.Rows
        vReports.DropDownItems.Add(vRow.Item("CriteriaSetDesc").ToString, Nothing, AddressOf MailingMenuHandler).Tag = vRow.Item("CriteriaSetNumber")
      Next
      If vReports.DropDownItems.Count = 0 Then vReports.Visible = False
      Return vReports.DropDownItems.Count
    End If
  End Function

  Private Sub MailingMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Dim vList As New ParameterList
      vList("CriteriaSet") = vMenuItem.Tag.ToString
      vList.IntegerValue("EventNumber") = mvEventInfo.EventNumber
      Dim vType As CareServices.TaskJobTypes
      Select Case EventDataType
        Case CareServices.XMLEventDataSelectionTypes.xedtEventBookings
          vType = CareServices.TaskJobTypes.tjtEventBookerMailing
        Case CareServices.XMLEventDataSelectionTypes.xedtEventAttendees
          vType = CareServices.TaskJobTypes.tjtEventDelegateMailing
        Case CareServices.XMLEventDataSelectionTypes.xedtEventPersonnel
          vType = CareServices.TaskJobTypes.tjtEventPersonnelMailing
        Case CareServices.XMLEventDataSelectionTypes.xedtEventCosts
          vType = CareServices.TaskJobTypes.tjtEventSponsorMailing
      End Select
      FormHelper.RunMailing(vType, vList)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
End Class

Public Class EventFinderMenu
  Inherits EventMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent)
  End Sub

  Protected Overrides Sub EventMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      mvMenuItems(EventMenuItems.emiDuplicateEvent).SetContextItemVisible(Me, True)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try

  End Sub

End Class