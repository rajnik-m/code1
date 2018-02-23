Friend Class FDEControlMenu
  Inherits ContextMenuStrip

  Friend Enum FDEControlmenuItems
    Customise
    Revert
    Parameters
    Delete
  End Enum

  Private mvUserControlName As String
  Private mvControlHasParameters As Boolean = False
  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Event MenuSelected(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FDEControlmenuItems)

  Public Sub New(ByVal pUserControlName As String, ByVal pControlHasParameters As Boolean)
    MyBase.New()

    With mvMenuItems
      .Add(FDEControlmenuItems.Customise.ToString, New MenuToolbarCommand(FDEControlmenuItems.Customise.ToString, ControlText.MnuFDECntrlCustomise, FDEControlmenuItems.Customise))
      .Add(FDEControlmenuItems.Revert.ToString, New MenuToolbarCommand(FDEControlmenuItems.Revert.ToString, ControlText.MnuFDECntrlRevert, FDEControlmenuItems.Revert))
      .Add(FDEControlmenuItems.Parameters.ToString, New MenuToolbarCommand(FDEControlmenuItems.Parameters.ToString, ControlText.MnuFDECntrlParameters, FDEControlmenuItems.Parameters))
      .Add(FDEControlmenuItems.Delete.ToString, New MenuToolbarCommand(FDEControlmenuItems.Delete.ToString, ControlText.MnuFDECntrlDelete, FDEControlmenuItems.Delete))
    End With

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next

    MenuToolbarCommand.SetAccessControl(mvMenuItems)

    mvUserControlName = pUserControlName
    mvControlHasParameters = pControlHasParameters

  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      Dim vMenu As MenuToolbarCommand = DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand)
      Dim vFDEControlMenuItem As FDEControlmenuItems = CType(vMenu.CommandID, FDEControlmenuItems)
      RaiseEvent MenuSelected(DirectCast(sender, ToolStripMenuItem), vFDEControlMenuItem)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub FDEControlMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    'mvMenuItems(FDEControlmenuItems.Customise).SetContextItemVisible(Me, (mvUserControlName <> "DISPLAYLABEL"))
    'mvMenuItems(FDEControlmenuItems.Revert).SetContextItemVisible(Me, (mvUserControlName <> "DISPLAYLABEL"))
    mvMenuItems(FDEControlmenuItems.Parameters).SetContextItemVisible(Me, (mvControlHasParameters = True))
  End Sub
End Class
