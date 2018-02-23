Imports System.Windows.Forms

Public Class frmCustomiseToolBar

  Private mvToolStrip As ToolStrip
  Private mvSeprator As Integer

  Public Sub New(ByVal pImg32 As ImageList, ByVal pImg16 As ImageList, ByVal pMenuItems As CollectionList(Of MenuToolbarCommand), ByVal pToolStrip As ToolStrip, ByVal pSeprator As Integer)
    mvSeprator = pSeprator
    InitializeComponent()
    InitialiseControls(pImg32, pImg16, pMenuItems, pToolStrip)
  End Sub

  Private Sub InitialiseControls(ByVal pImg32 As ImageList, ByVal pImg16 As ImageList, ByVal pMenuItems As CollectionList(Of MenuToolbarCommand), ByRef pToolStip As ToolStrip)
    Dim vMenuItemCollection As New CollectionList(Of MenuToolbarCommand)

    mvToolStrip = pToolStip
    SetControlTheme()
    Me.Text = ControlText.FrmCustomiseToolBar
    cmdOk.Text = ControlText.CmdOK
    cmdCancel.Text = ControlText.CmdCancel
    cmdApply.Text = ControlText.CmdApply
    cmdDefault.Text = ControlText.CmdDefaults

    lvwAvailable.SmallImageList = pImg16
    lvwAvailable.LargeImageList = pImg32
    lvwAvailable.Columns.Clear()
    lvwAvailable.Columns.Add("Command", lvwAvailable.Width)
    lvwAvailable.View = View.SmallIcon
    lvwAvailable.AllowDrop = True
    lvwAvailable.Sorting = SortOrder.Ascending
    lvwAvailable.LabelEdit = False  'BR17806

    lvwSelected.SmallImageList = pImg16
    lvwSelected.LargeImageList = pImg32
    lvwSelected.Columns.Clear()
    lvwSelected.Columns.Add("Command", lvwSelected.Width)
    lvwSelected.View = View.SmallIcon
    lvwSelected.AllowDrop = True

    For Each vMenuItems As MenuToolbarCommand In pMenuItems
      vMenuItemCollection.Add(vMenuItems.CommandID.ToString, vMenuItems)
    Next

    'Remove all the selected items from the available items list
    Dim vDuplicateItems As New CollectionList(Of MenuToolbarCommand)

    For Each vToolStripItem As ToolStripItem In pToolStip.Items
      Dim vCheckCommand1 As MenuToolbarCommand
      vCheckCommand1 = TryCast(vToolStripItem.Tag, MenuToolbarCommand)
      For Each vToolMenuItems As MenuToolbarCommand In vMenuItemCollection
        If (vToolMenuItems.CommandID = vCheckCommand1.CommandID AndAlso vCheckCommand1.CommandID <> mvSeprator) Then
          If (vDuplicateItems.Count > 0 And vDuplicateItems.ContainsKey(vCheckCommand1.CommandID.ToString)) Then
            Continue For
          End If
          vDuplicateItems.Add(vCheckCommand1.CommandID.ToString, vCheckCommand1)
        End If
      Next
    Next

    For Each vDuplicateItem As MenuToolbarCommand In vDuplicateItems
      vMenuItemCollection.Remove(vDuplicateItem)
    Next

    Dim vLVI As ListViewItem
    For Each vMenuItem As MenuToolbarCommand In vMenuItemCollection
      If vMenuItem.CommandID <= pImg32.Images.Count AndAlso vMenuItem.HideItem = False Then
        vLVI = New ListViewItem(vMenuItem.ToolBarText, vMenuItem.CommandID) 'BR17806
        vLVI.Tag = vMenuItem
        lvwAvailable.Items.Add(vLVI)
      End If
    Next
    lvwAvailable.Sort()
    Dim vLVI2 As ListViewItem
    Dim vCheckCommand As MenuToolbarCommand
    For Each vToolStripItem As ToolStripItem In pToolStip.Items
      vCheckCommand = TryCast(vToolStripItem.Tag, MenuToolbarCommand)
      If vCheckCommand.CommandID <= pImg32.Images.Count AndAlso vCheckCommand.HideItem = False Then
        vLVI2 = New ListViewItem(vCheckCommand.ToolBarText, vCheckCommand.CommandID)  'BR17806
        vLVI2.Tag = vCheckCommand
        lvwSelected.Items.Add(vLVI2)
      End If
    Next
    MainHelper.ToolbarTextPosition = DirectCast(mvToolStrip.Parent, IMainForm).MainMenu.ToolbarTextPosition
    If MainHelper.ToolbarTextPosition = TextImageRelation.ImageAboveText Then
      chkLabelsBelow.Checked = True
    Else
      chkLabelsBelow.Checked = False
    End If
  End Sub

  Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click
    If lvwAvailable.SelectedItems.Count > 0 Then
      Dim lvwItem As ListViewItem = lvwAvailable.SelectedItems(0)
      Dim vCommand As MenuToolbarCommand = CType(lvwItem.Tag, MenuToolbarCommand)
      If (vCommand.CommandID <> mvSeprator) Then
        'We found the command already exists in the toolbar so remove it
        lvwAvailable.SelectedItems(0).Remove()
      Else
        lvwAvailable.SelectedItems(0).Selected = False
      End If
      lvwSelected.Items.Add(DirectCast(lvwItem.Clone(), ListViewItem))
    End If
  End Sub

  Private Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemove.Click
    If lvwSelected.SelectedItems.Count > 0 Then
      Dim lvwItem As ListViewItem = lvwSelected.SelectedItems(0)
      lvwSelected.SelectedItems(0).Remove()
      Dim vCommand As MenuToolbarCommand = CType(lvwItem.Tag, MenuToolbarCommand)
      If vCommand.CommandID = mvSeprator Then
        If lvwAvailable.Items.Contains(lvwItem) Then lvwAvailable.Items.Add(lvwItem)
      Else
        lvwAvailable.Items.Add(lvwItem)
      End If

    End If
  End Sub

  Private Sub btnAddAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAll.Click
    For Each item As ListViewItem In lvwAvailable.Items
      lvwSelected.Items.Add(DirectCast(item.Clone(), ListViewItem))
    Next
    lvwAvailable.Clear()
    For Each vItem As ListViewItem In lvwSelected.Items
      Dim vCommand As MenuToolbarCommand = CType(vItem.Tag, MenuToolbarCommand)
      If (vCommand.CommandID = mvSeprator) Then
        lvwAvailable.Items.Add(DirectCast(vItem.Clone(), ListViewItem))
        Exit For
      End If
    Next
  End Sub

  Private Sub btnRemoveAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveAll.Click
    For Each vItem As ListViewItem In lvwSelected.Items
      Dim vCommand As MenuToolbarCommand = CType(vItem.Tag, MenuToolbarCommand)
      If (vCommand.CommandID = mvSeprator) Then
        If lvwAvailable.Items.Contains(vItem) Then lvwAvailable.Items.Add(DirectCast(vItem.Clone(), ListViewItem))
      Else
        lvwAvailable.Items.Add(DirectCast(vItem.Clone(), ListViewItem))
      End If
    Next
    lvwSelected.Clear()
  End Sub

  Private Function GetCurrentToolbarCommand() As MenuToolbarCommand
    Dim vCommand As MenuToolbarCommand = Nothing
    If lvwSelected.SelectedIndices.Count > 0 Then
      Dim vLVI As ListViewItem = lvwSelected.SelectedItems(0)
      If vLVI IsNot Nothing Then vCommand = CType(vLVI.Tag, MenuToolbarCommand)
    End If
    Return vCommand
  End Function
  Private Function GetCurrentToolbarCommandAvl() As MenuToolbarCommand
    Dim vCommand As MenuToolbarCommand = Nothing
    If lvwAvailable.SelectedIndices.Count > 0 Then
      Dim vLVI As ListViewItem = lvwAvailable.SelectedItems(0)
      If vLVI IsNot Nothing Then vCommand = CType(vLVI.Tag, MenuToolbarCommand)
    End If
    Return vCommand
  End Function
  Private Sub lvwSelected_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles lvwSelected.AfterLabelEdit
    Dim vMenuItem As MenuToolbarCommand = CType(lvwSelected.Items(e.Item).Tag, MenuToolbarCommand)
    'BR17806 - ToolBarText changes
    If e.Label Is Nothing Then
      vMenuItem.ToolBarText = txtLabel.Text
    Else
      vMenuItem.ToolBarText = e.Label
    End If
    txtLabel.Text = vMenuItem.ToolBarText
  End Sub

  Private Sub lvwSelected_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvwSelected.DragDrop, lvwAvailable.DragDrop

    If (lvwAvailable.SelectedItems.Count > 0) Then
      'Returns the location of the mouse pointer in the ListView control.
      Dim vPoint As Point = lvwSelected.PointToClient(New Point(e.X, e.Y))
      'Obtain the item that is located at the specified location of the mouse pointer.
      Dim dragAvailableToItem As ListViewItem = lvwSelected.GetItemAt(vPoint.X, vPoint.Y) 'lvwAvailable.GetItemAt(vPoint.X, vPoint.Y)
      'Obtain the index of the item at the mouse pointer.
      Dim vDragAvailableIndex As Integer
      'If the dragged item is nothing that means the item is dragged to the 
      'end of the list
      If dragAvailableToItem Is Nothing Then
        vDragAvailableIndex = lvwSelected.Items.Count + 1
      Else
        vDragAvailableIndex = dragAvailableToItem.Index
      End If

      Dim vCount As Integer
      Dim vSel(lvwAvailable.SelectedItems.Count) As ListViewItem
      For vCount = 0 To lvwAvailable.SelectedItems.Count - 1
        vSel(vCount) = lvwAvailable.SelectedItems.Item(vCount)
      Next
      For vCount = 0 To lvwAvailable.SelectedItems.Count - 1
        'Obtain the ListViewItem to be dragged to the target location.
        Dim vDragItem As ListViewItem = vSel(vCount)
        Dim vItemIndex As Integer = vDragAvailableIndex
        'If itemIndex = dragItem.Index Then Return
        If (vItemIndex > (lvwSelected.Items.Count)) Then
          vItemIndex = lvwSelected.Items.Count
        Else
          If vDragItem.Index < vItemIndex Then
            vItemIndex = vItemIndex + 1
          Else
            vItemIndex = vDragAvailableIndex + vCount
          End If
        End If
        'Insert the item in the specified location.
        Dim insertitem As ListViewItem = CType(vDragItem.Clone, ListViewItem)
        If Not lvwSelected.Items.Contains(insertitem) Then
          lvwSelected.Items.Insert(vItemIndex, insertitem)
        End If
        'Removes the item from the initial location while 
        'the item is moved to the new location.
        Dim vCommand As MenuToolbarCommand = CType(vDragItem.Tag, MenuToolbarCommand)
        If (vCommand.CommandID <> mvSeprator) Then
          'We found the command already exists in the toolbar so remove it
          lvwAvailable.Items.Remove(vDragItem)
        Else
          lvwAvailable.SelectedItems(0).Selected = False
        End If

      Next
      ReloadSelectedItem()
    Else
      'Return if the items are not selected in the ListView control.
      If lvwSelected.SelectedItems.Count = 0 Then Return
      'Returns the location of the mouse pointer in the ListView control.
      Dim vPoint As Point = lvwSelected.PointToClient(New Point(e.X, e.Y))
      'Obtain the item that is located at the specified location of the mouse pointer.
      Dim vDragToItem As ListViewItem = lvwSelected.GetItemAt(vPoint.X, vPoint.Y)
      If vDragToItem Is Nothing Then Return
      'Obtain the index of the item at the mouse pointer.
      Dim vDragIndex As Integer = vDragToItem.Index
      Dim vCount As Integer
      Dim vSel(lvwSelected.SelectedItems.Count) As ListViewItem
      For vCount = 0 To lvwSelected.SelectedItems.Count - 1
        vSel(vCount) = lvwSelected.SelectedItems.Item(vCount)
      Next
      For vCount = 0 To lvwSelected.SelectedItems.Count - 1
        'Obtain the ListViewItem to be dragged to the target location.
        Dim vDragItem As ListViewItem = vSel(vCount)
        Dim vItemIndex As Integer = vDragIndex
        If vItemIndex = vDragItem.Index Then Return
        If vDragItem.Index < vItemIndex Then
          vItemIndex = vItemIndex + 1
        Else
          vItemIndex = vDragIndex + vCount
        End If
        'Insert the item in the specified location.
        Dim insertitem As ListViewItem = CType(vDragItem.Clone, ListViewItem)
        lvwSelected.Items.Insert(vItemIndex, insertitem)
        'Removes the item from the initial location while 
        'the item is moved to the new location.
        lvwSelected.Items.Remove(vDragItem)
        ReloadSelectedItem()
      Next
    End If
  End Sub
  Private Sub ReloadSelectedItem()
    Dim vListViewItems(lvwSelected.Items.Count) As ListViewItem
    lvwSelected.Items.CopyTo(vListViewItems, 0)
    lvwSelected.Items.Clear()
    Dim i As Integer
    For i = 0 To (vListViewItems.GetUpperBound(0) - 1)
      lvwSelected.Items.Add(vListViewItems(i))
    Next
    lvwSelected.Refresh()
    lvwSelected.AutoArrange = False
  End Sub
  Private Sub lvwSelected_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvwSelected.DragEnter, lblAvailable.DragEnter
    Dim vCount As Integer
    For vCount = 0 To e.Data.GetFormats().Length - 1
      If e.Data.GetFormats()(vCount).Equals("System.Windows.Forms.ListView+SelectedListViewItemCollection") Then
        'The data from the drag source is moved to the target.
        e.Effect = DragDropEffects.Move
      End If
    Next
  End Sub

  Private Sub lvwSelected_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvwSelected.DragOver, lvwAvailable.DragOver
    If e.Data.GetDataPresent(GetType(MenuToolbarCommand).FullName) Then e.Effect = DragDropEffects.Copy
  End Sub

  Private Sub lvwSelected_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvwSelected.ItemDrag, lvwAvailable.ItemDrag
    Dim vLVI As ListViewItem = CType(e.Item, ListViewItem)
    If (lvwSelected.SelectedItems.Count > 0) Then
      lvwSelected.DoDragDrop(lvwSelected.SelectedItems, DragDropEffects.Move)
    Else
      lvwSelected.DoDragDrop(lvwAvailable.SelectedItems, DragDropEffects.Move)
    End If
  End Sub

  Private Sub lvwSelected_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvwSelected.SelectedIndexChanged
    Dim vMenuItem As MenuToolbarCommand = GetCurrentToolbarCommand()
    If vMenuItem IsNot Nothing Then
      txtToolTip.Text = vMenuItem.ToolTipText
      txtLabel.Text = vMenuItem.ToolBarText
    End If
  End Sub
  'BR17806
  Private Sub lvwAvailable_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvwAvailable.SelectedIndexChanged
    Dim vMenuItem As MenuToolbarCommand = GetCurrentToolbarCommandAvl()
    If vMenuItem IsNot Nothing Then
      txtToolTip.Text = vMenuItem.ToolTipText
      txtLabel.Text = vMenuItem.ToolBarText
    End If
  End Sub

  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDefault.Click
    MainHelper.ResetToolbar()
    'BR17806 - Added so that Selected List is refreshed with the Default List
    If mvToolStrip IsNot Nothing Then
      lvwSelected.Items.Clear()
      Dim vLVI2 As ListViewItem
      Dim vCheckCommand As MenuToolbarCommand
      For Each vToolStripItem As ToolStripItem In mvToolStrip.Items
        vCheckCommand = TryCast(vToolStripItem.Tag, MenuToolbarCommand)
        If vCheckCommand.CommandID <= lvwSelected.LargeImageList.Images.Count AndAlso vCheckCommand.HideItem = False Then
          vLVI2 = New ListViewItem(vCheckCommand.ToolBarText, vCheckCommand.CommandID)  'BR17806
          vLVI2.Tag = vCheckCommand
          lvwSelected.Items.Add(vLVI2)
        End If
      Next
    End If
  End Sub

  Private Sub cmdApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdApply.Click
    RebuildToolBar()
  End Sub

  Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    RebuildToolBar()
    Me.Close()
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    DirectCast(mvToolStrip.Parent, IMainForm).MainMenu.ResetToolbar()
    Me.Close()
  End Sub

  Private Sub RebuildToolBar()
    'Clear the toolstrip and add form 
    mvToolStrip.Items.Clear()

    For Each vListItem As ListViewItem In lvwSelected.Items
      Dim vCheckCommand As MenuToolbarCommand = TryCast(vListItem.Tag, MenuToolbarCommand)

      Dim vToolStripItem As ToolStripItem = TryCast(vCheckCommand.MenuStripItem, ToolStripItem)
      If vToolStripItem IsNot Nothing Then
        vToolStripItem = vCheckCommand.AddToToolStrip(mvToolStrip)
      End If

      If vToolStripItem IsNot Nothing Then
        If (chkLabelsBelow.Checked) Then
          vToolStripItem.TextImageRelation = TextImageRelation.ImageAboveText
        End If
        If vToolStripItem.Text.Length > 0 Then
          vToolStripItem.AccessibleName = vToolStripItem.Text
        Else
          vToolStripItem.AccessibleName = vToolStripItem.ToolTipText
        End If
      End If
    Next
    DirectCast(mvToolStrip.Parent, IMainForm).MainMenu.SaveToolbarItems()

  End Sub

  Private Function GetDragCursor(ByVal pText As String, ByVal pFont As Font) As Cursor
    Dim vBmp As New Bitmap(1, 1)
    Dim g As Graphics = Graphics.FromImage(vBmp)
    Dim sz As SizeF = g.MeasureString(pText, pFont)
    vBmp = New Bitmap(CInt(sz.Width), CInt(sz.Height))
    g = Graphics.FromImage(vBmp)
    g.Clear(Color.White)
    g.DrawString(pText, pFont, Brushes.Black, 0, 0)
    g.Dispose()
    Dim vCursor As New Cursor(vBmp.GetHicon)
    vBmp.Dispose()
    Return vCursor
  End Function

  Private Sub frmCustomiseToolBar_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    lvwSelected.SmallImageList = Nothing
    lvwSelected.LargeImageList = Nothing
  End Sub

  Private Sub frmCustomiseToolBar_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Location = MDILocation(Width, Height)
    bpl1.RepositionButtons()
    bpl2.RepositionButtons()
  End Sub

  Private Sub chkLabelsBelow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLabelsBelow.Click
    If chkLabelsBelow.Checked Then
      MainHelper.ToolbarTextPosition = TextImageRelation.ImageAboveText
    Else
      MainHelper.ToolbarTextPosition = TextImageRelation.ImageBeforeText
    End If
  End Sub

  Private Sub txtLabel_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLabel.LostFocus
    Dim vMenuItem As MenuToolbarCommand = GetCurrentToolbarCommand()
    If vMenuItem IsNot Nothing Then vMenuItem.ToolBarText = txtLabel.Text
  End Sub

  Private Sub txtToolTip_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtToolTip.LostFocus
    Dim vCommand As MenuToolbarCommand = Nothing
    If lvwSelected.SelectedIndices.Count > 0 Then
      Dim vLVI As ListViewItem = lvwSelected.SelectedItems(0)
      If vLVI IsNot Nothing Then vCommand = CType(vLVI.Tag, MenuToolbarCommand)
      If vCommand IsNot Nothing Then
        vCommand.ToolTipText = txtToolTip.Text
        vLVI.Text = txtLabel.Text
      End If
    End If
  End Sub

End Class
