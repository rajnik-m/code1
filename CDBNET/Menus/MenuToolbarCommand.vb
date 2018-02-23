Public Class MenuToolbarCommand

  Private Enum MenuToolbarCommandType
    mtctButton
    mtctSeparator
  End Enum

  Private mvMenuText As String
  Private mvCommandID As Integer
  Private mvCommandName As String
  Private mvMenuImage As Image
  Private mvAccessControlID As String
  Private mvToolTipText As String         'Toolbar Tooltip
  Private mvImageIndex As Integer         'Toolbar ImageIndex
  Private mvOnClick As EventHandler
  Private mvHideItem As Boolean
  Private mvType As MenuToolbarCommandType
  Private mvToolStrip As ToolStrip
  Private mvEntityGroup As EntityGroup

  Public Shared Function NewMenuItem(ByVal pItem As Integer, ByVal pText As String, ByVal pOnClick As System.EventHandler) As ToolStripMenuItem
    Return NewMenuItem(pItem, pText, Nothing, pOnClick)
  End Function
  Public Shared Function NewMenuItem(ByVal pItem As Integer, ByVal pText As String, Optional ByVal pImage As Image = Nothing, Optional ByVal pOnClick As System.EventHandler = Nothing) As ToolStripMenuItem
    Dim vItem As New ToolStripMenuItem(pText, pImage, pOnClick)
    vItem.Tag = pItem
    Return vItem
  End Function

  Public Shared Sub SetAccessControl(ByVal pMenuItems As CollectionList(Of MenuToolbarCommand), Optional ByVal pMainMenu As Boolean = False)
    Dim vMenuTable As DataTable = DataHelper.GetCachedLookupData(CareServices.XMLLookupDataTypes.xldtUserMenuAccess)
    If Not vMenuTable Is Nothing Then
      Dim vItemID As String
      For Each vRow As DataRow In vMenuTable.Rows
        vItemID = vRow.Item("Item").ToString
        If vItemID.StartsWith("SC") OrElse pMainMenu = False Then
          If vRow.Item("Visible").ToString = "N" Then
            For Each vItem As MenuToolbarCommand In pMenuItems
              If vItem.AccessControlID = vItemID Then
                vItem.HideItem = True
                Exit For
              End If
            Next
          End If
        Else
          If pMainMenu Then AppValues.AddAccessControlItem(vItemID, vRow.Item("Visible").ToString = "Y")
        End If
      Next
    End If
  End Sub

  Public Sub New(ByVal pCommandName As String, ByVal pCommandID As Integer)
    mvCommandName = pCommandName
    mvToolTipText = pCommandName
    mvCommandID = pCommandID
    mvType = MenuToolbarCommandType.mtctSeparator
  End Sub
  Public Sub New(ByVal pCommandName As String, ByVal pMenuText As String, ByVal pCommandID As Integer, Optional ByVal pAccessControlID As String = "", Optional ByVal pImage As Image = Nothing, Optional ByVal pToolTip As String = "")
    mvCommandName = pCommandName
    mvMenuText = pMenuText
    mvCommandID = pCommandID
    mvAccessControlID = pAccessControlID
    mvMenuImage = pImage
    mvToolTipText = pToolTip
    mvType = MenuToolbarCommandType.mtctButton
  End Sub

  Public Function AddToMenu(ByVal pMenu As ToolStripMenuItem, Optional ByVal pChecked As Boolean = False) As ToolStripMenuItem
    Dim vItem As ToolStripMenuItem = Nothing
    If Not mvHideItem Then
      vItem = MenuStripItem(pChecked)
      pMenu.DropDownItems.Add(vItem)
    End If
    Return vItem
  End Function
  Public Function MenuStripItem(Optional ByVal pChecked As Boolean = False) As ToolStripMenuItem
    Dim vItem As New ToolStripMenuItem(mvMenuText, mvMenuImage, mvOnClick)
    vItem.Name = "msi" & mvCommandName
    vItem.Tag = Me
    If pChecked Then vItem.CheckState = CheckState.Checked
    Return vItem
  End Function
  Public Function AddToToolStrip(ByVal pToolStrip As ToolStrip, Optional ByVal pChecked As Boolean = False) As ToolStripItem
    Dim vItem As ToolStripItem = Nothing
    If Not mvHideItem Then
      If mvType = MenuToolbarCommandType.mtctButton Then
        vItem = ToolStripButton(pChecked)
      Else
        vItem = ToolStripSeparator()
      End If
      mvToolStrip = pToolStrip
      pToolStrip.Items.Add(vItem)
    End If
    Return vItem
  End Function
  Public Function AddToToolStripAt(ByVal pToolStrip As ToolStrip, ByVal pIndex As Integer, Optional ByVal pChecked As Boolean = False) As ToolStripItem
    Dim vItem As ToolStripItem = Nothing
    If Not mvHideItem Then
      If mvType = MenuToolbarCommandType.mtctButton Then
        vItem = ToolStripButton(pChecked)
      Else
        vItem = ToolStripSeparator()
      End If
      mvToolStrip = pToolStrip
      pToolStrip.Items.Insert(pIndex, vItem)
    End If
    Return vItem
  End Function
  Public Function ToolStripButton(Optional ByVal pChecked As Boolean = False) As ToolStripButton
    Dim vButton As New ToolStripButton
    vButton.ImageIndex = mvCommandID
    vButton.ToolTipText = mvToolTipText
    AddHandler vButton.Click, mvOnClick
    AddHandler vButton.MouseDown, AddressOf ButtonMouseDown
    AddHandler vButton.MouseUp, AddressOf ButtonMouseUp
    AddHandler vButton.MouseMove, AddressOf ButtonMouseMove
    vButton.Tag = Me
    vButton.Name = "tsb" & mvCommandName
    If pChecked Then vButton.Checked = True
    Return vButton
  End Function
  Public Function ToolStripSeparator() As ToolStripSeparator
    Dim vItem As New ToolStripSeparator
    vItem.ImageIndex = mvCommandID
    'vItem.ToolTipText = mvCommandName
    AddHandler vItem.MouseDown, AddressOf ButtonMouseDown
    AddHandler vItem.MouseUp, AddressOf ButtonMouseUp
    AddHandler vItem.MouseMove, AddressOf ButtonMouseMove
    vItem.Tag = Me
    Return vItem
  End Function

  Private mvMouseDown As Boolean
  Private mvMouseLocation As Point

  Private Sub ButtonMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
    If e.Button = MouseButtons.Left Then
      mvMouseDown = True
      mvMouseLocation = e.Location
    End If
  End Sub
  Private Sub ButtonMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
    If e.Button = MouseButtons.Left Then mvMouseDown = False
  End Sub
  Private Sub ButtonMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
    If mvMouseDown Then
      Dim vXOffset As Integer = Math.Abs(e.Location.X - mvMouseLocation.X)
      Dim vYOffset As Integer = Math.Abs(e.Location.Y - mvMouseLocation.Y)
      If ((vXOffset > 4) And (mvType = MenuToolbarCommandType.mtctButton)) OrElse _
         ((vXOffset > 1) And (mvType = MenuToolbarCommandType.mtctSeparator)) OrElse vYOffset > 4 Then
        mvMouseDown = False
        Dim vItem As ToolStripItem = CType(sender, ToolStripItem)
        If vItem.DoDragDrop(vItem.Tag, DragDropEffects.Copy) = DragDropEffects.Copy Then
          mvToolStrip.Items.Remove(vItem)
        End If
      End If
    End If
  End Sub

  Public WriteOnly Property OnClick() As EventHandler
    Set(ByVal pValue As EventHandler)
      mvOnClick = pValue
    End Set
  End Property
  Public ReadOnly Property CommandID() As Integer
    Get
      Return mvCommandID
    End Get
  End Property
  Public ReadOnly Property AccessControlID() As String
    Get
      Return mvAccessControlID
    End Get
  End Property
  Public Property ToolTipText() As String
    Get
      Return mvToolTipText
    End Get
    Set(ByVal pValue As String)
      mvToolTipText = pValue
    End Set
  End Property
  Public Property MenuText() As String
    Get
      Return mvMenuText
    End Get
    Set(ByVal pValue As String)
      mvMenuText = pValue
    End Set
  End Property
  Public Property EntityGroup() As EntityGroup
    Get
      Return mvEntityGroup
    End Get
    Set(ByVal pValue As EntityGroup)
      mvEntityGroup = pValue
    End Set
  End Property

  Public Property HideItem() As Boolean
    Get
      Return mvHideItem
    End Get
    Set(ByVal pValue As Boolean)
      mvHideItem = pValue
    End Set
  End Property

  Public Sub SetContextItemVisible(ByVal pContextMenu As ContextMenuStrip, ByVal pVisible As Boolean)
    If mvHideItem Then pVisible = False
    For Each vItem As ToolStripMenuItem In pContextMenu.Items
      If vItem.Name = "msi" & mvCommandName Then
        vItem.Visible = pVisible
        Exit For
      End If
    Next
  End Sub

  Public Sub SetContextItemEnabled(ByVal pContextMenu As ContextMenuStrip, ByVal pEnabled As Boolean)
    If mvHideItem = False Then
      For Each vItem As ToolStripMenuItem In pContextMenu.Items
        If vItem.Name = "msi" & mvCommandName Then
          vItem.Enabled = pEnabled
          Exit For
        End If
      Next
    End If
  End Sub

  Public Sub EnableToolStripItem(ByVal pMenuStrip As MenuStrip, ByVal pToolStrip As ToolStrip, ByVal pEnable As Boolean)
    Dim vItem As ToolStripMenuItem = FindMenuStripItem(pMenuStrip)
    If Not vItem Is Nothing Then vItem.Enabled = pEnable
    If Not pToolStrip Is Nothing Then
      Dim vButton As ToolStripButton = FindToolStripItem(pToolStrip)
      If Not vButton Is Nothing Then vButton.Enabled = pEnable
    End If
  End Sub
  Public Sub CheckToolStripItem(ByVal pMenuStrip As MenuStrip, ByVal pToolStrip As ToolStrip, ByVal pChecked As Boolean)
    Dim vItem As ToolStripMenuItem = FindMenuStripItem(pMenuStrip)
    If Not vItem Is Nothing Then vItem.Checked = pChecked
    If Not pToolStrip Is Nothing Then
      Dim vButton As ToolStripButton = FindToolStripItem(pToolStrip)
      If Not vButton Is Nothing Then vButton.Checked = pChecked
    End If
  End Sub
  Public Function FindToolStripItem(ByVal pToolStrip As ToolStrip) As ToolStripButton
    For Each vItem As ToolStripItem In pToolStrip.Items
      If TypeOf vItem Is ToolStripButton Then
        Dim vCommand As MenuToolbarCommand = DirectCast(vItem.Tag, MenuToolbarCommand)
        If vCommand.CommandID = mvCommandID Then
          Return CType(vItem, ToolStripButton)
        End If
      End If
    Next
    Debug.Assert(True, "FindToolStripItem failed to find " & mvCommandName)
    Return Nothing
  End Function
  Public Function FindMenuStripItem(ByVal pMenuStrip As MenuStrip) As ToolStripMenuItem
    For Each vItem As ToolStripMenuItem In pMenuStrip.Items
      If vItem.DropDownItems.ContainsKey("msi" & mvCommandName) Then
        Return CType(vItem.DropDownItems("msi" & mvCommandName), ToolStripMenuItem)
      End If
    Next
    Debug.Assert(True, "FindMenuStripItem failed to find " & mvCommandName)
    Return Nothing
  End Function

  Public Function FindMenuStripItem(ByVal pContextMenu As ContextMenuStrip) As ToolStripMenuItem
    For Each vItem As ToolStripMenuItem In pContextMenu.Items
      If vItem.Name = "msi" & mvCommandName Then
        Return vItem
      End If
    Next
    Debug.Assert(True, "FindMenuStripItem failed to find " & mvCommandName)
    Return Nothing
  End Function

End Class
