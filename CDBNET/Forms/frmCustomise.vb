Imports System.Windows.Forms

Public Class frmCustomise

  Public Sub New(ByVal pImg32 As ImageList, ByVal pImg16 As ImageList, ByVal pMenuItems As CollectionList(Of MenuToolbarCommand))
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pImg32, pImg16, pMenuItems)
  End Sub

  Private Sub InitialiseControls(ByVal pImg32 As ImageList, ByVal pImg16 As ImageList, ByVal pMenuItems As CollectionList(Of MenuToolbarCommand))
    SetControlTheme()
    Me.Text = ControlText.FrmCustomise
    lvw.SmallImageList = pImg16
    lvw.LargeImageList = pImg32
    lvw.Columns.Clear()
    lvw.Columns.Add("Command", lvw.Width)
    lvw.View = View.SmallIcon
    lvw.AllowDrop = True

    Dim vLVI As ListViewItem
    For Each vMenuItem As MenuToolbarCommand In pMenuItems
      If vMenuItem.CommandID < pImg32.Images.Count AndAlso vMenuItem.HideItem = False Then
        vLVI = New ListViewItem(vMenuItem.ToolTipText, vMenuItem.CommandID)
        vLVI.Tag = vMenuItem
        lvw.Items.Add(vLVI)
      End If
    Next
    lvw.Sorting = SortOrder.Ascending
    lvw.Sort()
    chkLabelsBelow.Checked = MainHelper.ToolbarTextPosition = TextImageRelation.ImageAboveText
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub lvw_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles lvw.AfterLabelEdit
    Dim vMenuItem As MenuToolbarCommand = CType(lvw.Items(e.Item).Tag, MenuToolbarCommand)
    vMenuItem.ToolTipText = e.Label
    txtToolTip.Text = vMenuItem.ToolTipText
  End Sub

  Private Sub lvw_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvw.DragDrop
    If e.Data.GetDataPresent(GetType(MenuToolbarCommand).FullName) Then e.Effect = DragDropEffects.Copy
  End Sub

  Private Sub lvw_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvw.DragOver
    If e.Data.GetDataPresent(GetType(MenuToolbarCommand).FullName) Then e.Effect = DragDropEffects.Copy
  End Sub

  Private Sub lvw_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvw.ItemDrag
    Dim vLVI As ListViewItem = CType(e.Item, ListViewItem)
    DoDragDrop(vLVI.Tag, DragDropEffects.Copy)
  End Sub

  Private Sub frmCustomise_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    'lvw.Items.Clear()
    lvw.SmallImageList = Nothing
    lvw.LargeImageList = Nothing
    'lvw.Dispose()
    'bpl.Controls.Clear()
    'bpl.Dispose()
  End Sub

  Private Sub frmCustomise_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Location = MDILocation(Width, Height)
    bpl.RepositionButtons()
  End Sub

  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
    MainHelper.ResetToolbar()
  End Sub

  Private Sub lvw_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvw.SelectedIndexChanged
    Dim vMenuItem As MenuToolbarCommand = GetCurrentToolbarCommand()
    If vMenuItem IsNot Nothing Then
      txtToolTip.Text = vMenuItem.ToolTipText
      txtLabel.Text = vMenuItem.ToolBarText
    End If
  End Sub

  Private Sub txtToolTip_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtToolTip.LostFocus
    Dim vCommand As MenuToolbarCommand = Nothing
    If lvw.SelectedIndices.Count > 0 Then
      Dim vLVI As ListViewItem = lvw.SelectedItems(0)
      If vLVI IsNot Nothing Then vCommand = CType(vLVI.Tag, MenuToolbarCommand)
      If vCommand IsNot Nothing Then
        vCommand.ToolTipText = txtToolTip.Text
        vLVI.Text = txtToolTip.Text
      End If
    End If
  End Sub

  Private Sub txtLabel_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLabel.LostFocus
    Dim vMenuItem As MenuToolbarCommand = GetCurrentToolbarCommand()
    If vMenuItem IsNot Nothing Then vMenuItem.ToolBarText = txtLabel.Text
  End Sub

  Private Function GetCurrentToolbarCommand() As MenuToolbarCommand
    Dim vCommand As MenuToolbarCommand = Nothing
    If lvw.SelectedIndices.Count > 0 Then
      Dim vLVI As ListViewItem = lvw.SelectedItems(0)
      If vLVI IsNot Nothing Then vCommand = CType(vLVI.Tag, MenuToolbarCommand)
    End If
    Return vCommand
  End Function

  Private Sub chkLabelsBelow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLabelsBelow.Click
    If chkLabelsBelow.Checked Then
      MainHelper.ToolbarTextPosition = TextImageRelation.ImageAboveText
    Else
      MainHelper.ToolbarTextPosition = TextImageRelation.ImageBeforeText
    End If
  End Sub
End Class
