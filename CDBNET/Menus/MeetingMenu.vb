Imports System.IO

Public Class MeetingMenu
  Inherits ContextMenuStrip

  Private mvParent As MaintenanceParentForm
  Private mvMeetingNumber As Integer

  Public Event MenuSelected(ByVal pItem As MeetingMenuItems)

  Public Enum MeetingMenuItems
    mmiNew
    mmiEdit
    mmiDelete
    mniDuplicateMeeting
  End Enum

  Protected mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent

    With mvMenuItems
      .Add(MeetingMenuItems.mmiNew.ToString, New MenuToolbarCommand(MeetingMenuItems.mmiNew.ToString, ControlText.MnuMeetingNew, MeetingMenuItems.mmiNew, "CDDPMN"))
      .Add(MeetingMenuItems.mmiEdit.ToString, New MenuToolbarCommand(MeetingMenuItems.mmiEdit.ToString, ControlText.MnuMeetingEdit, MeetingMenuItems.mmiEdit, "CDDPUM"))
      .Add(MeetingMenuItems.mmiDelete.ToString, New MenuToolbarCommand(MeetingMenuItems.mmiDelete.ToString, ControlText.MnuMeetingDelete, MeetingMenuItems.mmiDelete, "CDDPDM"))
      .Add(MeetingMenuItems.mniDuplicateMeeting.ToString, New MenuToolbarCommand(MeetingMenuItems.mniDuplicateMeeting.ToString, ControlText.MnuMeetingDuplicate, MeetingMenuItems.mniDuplicateMeeting, ""))
    End With
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      vItem.MenuStripItem.Visible = vItem.HideItem = False
      Me.Items.Add(vItem.MenuStripItem)
    Next
  End Sub

  Public Property MeetingNumber() As Integer
    Get
      Return mvMeetingNumber
    End Get
    Set(ByVal Value As Integer)
      mvMeetingNumber = Value
    End Set
  End Property

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As MeetingMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, MeetingMenuItems)

      Select Case vMenuItem
        Case MeetingMenuItems.mmiNew, MeetingMenuItems.mmiEdit
          Dim vMeetingToEdit As Integer = 0
          If vMenuItem = MeetingMenuItems.mmiEdit Then vMeetingToEdit = mvMeetingNumber
          FormHelper.EditMeeting(vMeetingToEdit, mvParent)
        Case MeetingMenuItems.mmiDelete
          If Not ConfirmDelete() Then Exit Sub
          Dim vList As ParameterList = New ParameterList(True)
          vList.IntegerValue("MeetingNumber") = mvMeetingNumber
          DataHelper.DeleteItem(CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetings, CareServices.XMLMaintenanceControlTypes), vList)
          mvParent.RefreshData(CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetings, CareServices.XMLMaintenanceControlTypes))
        Case MeetingMenuItems.mniDuplicateMeeting
          RaiseEvent MenuSelected(MeetingMenuItems.mniDuplicateMeeting)
          'Dim vList As ParameterList = New ParameterList(True)
          'vList.IntegerValue("MeetingNumber") = mvMeetingNumber
          'Dim vMeetingNumber As Integer = 0
          'Dim vDefaults As New ParameterList
          'vDefaults("Description") = "blah"  'need to set meeting desc and time to default from meeting
          'Dim vList1 As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateMeeting, vDefaults)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
End Class


Public Class MeetingFinderMenu
  Inherits MeetingMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent)
  End Sub

  Protected Sub MeetingMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      For Each vItem As ToolStripItem In Me.Items
        vItem.Visible = False
      Next
      mvMenuItems(MeetingMenuItems.mniDuplicateMeeting).SetContextItemVisible(Me, True)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try

  End Sub

End Class