Public Class ServiceBookingMenu
  Inherits ContextMenuStrip

  Private mvParent As MaintenanceParentForm
  Private mvSelectionSetNumber As Integer
  Protected mvDataRow As DataRow
  Private mvChangePPDs As Boolean

  Public Event MenuSelected(ByVal pItem As ServiceBookingMenuItems, ByVal pDataRow As DataRow, ByVal pChangeDetails As Boolean)

  Public Enum ServiceBookingMenuItems
    sbiCancel
    sbiGoToServiceContact
    sbiSBGoToRelatedContact
  End Enum

  Public Sub SetContext(ByVal pDST As CareServices.XMLContactDataSelectionTypes, ByVal pDataRow As DataRow)
    mvDataRow = pDataRow
  End Sub

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent
    With mvMenuItems
      .Add(ServiceBookingMenuItems.sbiCancel.ToString, New MenuToolbarCommand(ServiceBookingMenuItems.sbiCancel.ToString, ControlText.mnuServiceBookingCancel, ServiceBookingMenuItems.sbiCancel, ""))
      .Add(ServiceBookingMenuItems.sbiGoToServiceContact.ToString, New MenuToolbarCommand(ServiceBookingMenuItems.sbiGoToServiceContact.ToString, ControlText.MnuServiceBookingGoToServiceContact, ServiceBookingMenuItems.sbiGoToServiceContact, ""))
      .Add(ServiceBookingMenuItems.sbiSBGoToRelatedContact.ToString, New MenuToolbarCommand(ServiceBookingMenuItems.sbiSBGoToRelatedContact.ToString, ControlText.MnuServiceBookingGoToRelatedContact, ServiceBookingMenuItems.sbiSBGoToRelatedContact, ""))
    End With
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MainMenuHandler
      vItem.MenuStripItem.Visible = vItem.HideItem = False
      Me.Items.Add(vItem.MenuStripItem)
    Next
  End Sub

  Public Property SelectionSetNumber() As Integer
    Get
      Return mvSelectionSetNumber
    End Get
    Set(ByVal Value As Integer)
      mvSelectionSetNumber = Value
    End Set
  End Property

  Private Sub MainMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MenuHandler(DirectCast(sender, ToolStripMenuItem), CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, ServiceBookingMenuItems))
  End Sub

  Protected Sub MenuHandler(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As ServiceBookingMenuItems)
    Dim vCursor As New BusyCursor
    Try
      RaiseEvent MenuSelected(pItem, mvDataRow, mvChangePPDs)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub


  Private Sub ServiceBookingMenu_opeing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    Try
      'Setting Cancel option
      If mvDataRow IsNot Nothing Then
        If mvDataRow("CancellationReason").ToString.Length = 0 Then
          Me.Items(ServiceBookingMenuItems.sbiCancel).Visible = DateDiff(DateInterval.Day, DateValue(AppValues.TodaysDate()), DateValue(mvDataRow("EndDate").ToString)) >= 0
        Else
          Me.Items(ServiceBookingMenuItems.sbiCancel).Visible = False
        End If
        'Setting GoToRelatedContact option
        If IntegerValue(mvDataRow("RelatedContactNumber").ToString) > 0 Then
          Me.Items(ServiceBookingMenuItems.sbiSBGoToRelatedContact).Visible = True
        Else
          Me.Items(ServiceBookingMenuItems.sbiSBGoToRelatedContact).Visible = False
        End If
        'Setting GoToServiceContact option
        Me.Items(ServiceBookingMenuItems.sbiGoToServiceContact).Visible = True
      Else
        Me.Items(ServiceBookingMenuItems.sbiCancel).Visible = False
        Me.Items(ServiceBookingMenuItems.sbiSBGoToRelatedContact).Visible = False
        Me.Items(ServiceBookingMenuItems.sbiGoToServiceContact).Visible = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub




End Class
