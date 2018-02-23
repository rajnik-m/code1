Public Class TaskBarIconMenu
  Inherits ContextMenuStrip

#Region "Enums and Module Variables"
  Public Enum TaskbarIconMenuItems
    tbmiShowNotification
    tbmiJobStatus
  End Enum

  Private mvShowNotification As Boolean
  Private mvShowTask As Boolean

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)
#End Region

  Public Event ShowNotifications(ByVal pSender As Object, ByVal pName As String)
  Public Event ShowTaskStatus(ByVal pSender As Object, ByVal pName As String)

#Region "Property"
  Public Property ShowTaskItems As Boolean
    Get
      Return mvShowTask
    End Get
    Set(ByVal value As Boolean)
      mvShowTask = value
    End Set
  End Property

  Public Property ShowNotificationItems As Boolean
    Get
      Return mvShowNotification
    End Get
    Set(ByVal value As Boolean)
      mvShowNotification = value
    End Set
  End Property
#End Region

#Region "Constructor"
  Public Sub New()
    MyBase.New()

    With mvMenuItems
      .Add(TaskbarIconMenuItems.tbmiShowNotification.ToString, New MenuToolbarCommand("ShowNotification", ControlText.MnuNotification, TaskbarIconMenuItems.tbmiShowNotification))
      .Add(TaskbarIconMenuItems.tbmiJobStatus.ToString, New MenuToolbarCommand("ShowActiveJobs", ControlText.MnuActiveTask, TaskbarIconMenuItems.tbmiJobStatus))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
  End Sub
#End Region
#Region "Event Handlers"
  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vMenuItem As TaskbarIconMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, TaskbarIconMenuItems)

    Select Case vMenuItem
      Case TaskbarIconMenuItems.tbmiShowNotification
        RaiseEvent ShowNotifications(sender, "Notification")
      Case TaskbarIconMenuItems.tbmiJobStatus
        RaiseEvent ShowTaskStatus(sender, "Task")
      Case Else
    End Select
  End Sub
#End Region

  Private Sub TaskBarIconMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    Dim vCursor As New BusyCursor
    Try
      Me.Items(TaskbarIconMenuItems.tbmiShowNotification).Visible = mvShowNotification
      Me.Items(TaskbarIconMenuItems.tbmiJobStatus).Visible = mvShowTask
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
End Class
