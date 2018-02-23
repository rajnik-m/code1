Public Class WebPopupMenu
  Inherits ContextMenuStrip

  Public Enum WebMenuItems
    wmiNewPage
    wmiNewPageItem
    wmiNewMenu
    wmiNewMenuItem
    wmiNewImage
    wmiNewDocument
    wmiCopyPage
    wmiDeletePage
    wmiDeletePageItem
    wmiDeleteMenu
    wmiDeleteMenuItem
    wmiRefresh
    wmiAddPage
    wmiExport
    wmiImport
    wmiMoveItemUp
    wmiMoveItemDown
  End Enum

  Private mvDataType As CareWebAccess.XMLWebDataSelectionTypes
  Private mvID As Integer
  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Event MenuSelected(ByVal pItem As WebMenuItems)

  Public Sub New()
    MyBase.New()

    With mvMenuItems
      .Add(WebMenuItems.wmiNewPage.ToString, New MenuToolbarCommand("NewPage", "New Page", WebMenuItems.wmiNewPage))
      .Add(WebMenuItems.wmiNewPageItem.ToString, New MenuToolbarCommand("NewPageItem", "New Page Item", WebMenuItems.wmiNewPageItem))
      .Add(WebMenuItems.wmiNewMenu.ToString, New MenuToolbarCommand("NewMenu", "New Menu", WebMenuItems.wmiNewMenu))
      .Add(WebMenuItems.wmiNewMenuItem.ToString, New MenuToolbarCommand("NewMenuItem", "New Menu Item", WebMenuItems.wmiNewMenuItem))
      .Add(WebMenuItems.wmiNewImage.ToString, New MenuToolbarCommand("NewImage", "New Image", WebMenuItems.wmiNewImage))
      .Add(WebMenuItems.wmiNewDocument.ToString, New MenuToolbarCommand("NewDocument", "New Document", WebMenuItems.wmiNewDocument))
      .Add(WebMenuItems.wmiCopyPage.ToString, New MenuToolbarCommand("CopyPage", "Copy Page", WebMenuItems.wmiCopyPage))
      .Add(WebMenuItems.wmiDeletePage.ToString, New MenuToolbarCommand("DeletePage", "Delete Page", WebMenuItems.wmiDeletePage))
      .Add(WebMenuItems.wmiDeletePageItem.ToString, New MenuToolbarCommand("DeletePageItem", "Delete Page Item", WebMenuItems.wmiDeletePageItem))
      .Add(WebMenuItems.wmiDeleteMenu.ToString, New MenuToolbarCommand("DeleteMenu", "Delete Menu", WebMenuItems.wmiDeleteMenu))
      .Add(WebMenuItems.wmiDeleteMenuItem.ToString, New MenuToolbarCommand("DeleteMenuItem", "Delete Menu Item", WebMenuItems.wmiDeleteMenuItem))
      .Add(WebMenuItems.wmiRefresh.ToString, New MenuToolbarCommand("Refresh", "Refresh", WebMenuItems.wmiRefresh))
      .Add(WebMenuItems.wmiAddPage.ToString, New MenuToolbarCommand("AddPage", "Add Page", WebMenuItems.wmiAddPage))
      .Add(WebMenuItems.wmiExport.ToString, New MenuToolbarCommand("Export", "Export Web", WebMenuItems.wmiExport))
      .Add(WebMenuItems.wmiImport.ToString, New MenuToolbarCommand("Import", "Import Web", WebMenuItems.wmiImport))
      .Add(WebMenuItems.wmiMoveItemUp.ToString, New MenuToolbarCommand("MoveItemUp", "Move Item Up", WebMenuItems.wmiMoveItemUp))
      .Add(WebMenuItems.wmiMoveItemDown.ToString, New MenuToolbarCommand("MoveItemDown", "Move Item Down", WebMenuItems.wmiMoveItemDown))
    End With

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
  End Sub

  Public Sub SetContext(ByVal pType As CareWebAccess.XMLWebDataSelectionTypes, ByVal pID As Integer)
    mvDataType = pType
    mvID = pID
  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As WebMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, WebMenuItems)
      Select Case vMenuItem
        Case WebMenuItems.wmiNewPage, WebMenuItems.wmiNewPageItem, WebMenuItems.wmiNewMenu, WebMenuItems.wmiNewMenuItem, WebMenuItems.wmiNewImage, _
             WebMenuItems.wmiDeletePage, WebMenuItems.wmiDeleteMenu, WebMenuItems.wmiDeleteMenuItem, WebMenuItems.wmiDeletePageItem, _
             WebMenuItems.wmiRefresh, WebMenuItems.wmiAddPage, WebMenuItems.wmiNewDocument, WebMenuItems.wmiCopyPage, _
             WebMenuItems.wmiExport, WebMenuItems.wmiImport, WebMenuItems.wmiMoveItemUp, WebMenuItems.wmiMoveItemDown
          RaiseEvent MenuSelected(vMenuItem)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub WebPopupMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    For Each vItem As ToolStripItem In Me.Items
      vItem.Visible = False
    Next
    Select Case mvDataType
      Case CareWebAccess.XMLWebDataSelectionTypes.wstControl
        mvMenuItems(WebMenuItems.wmiRefresh).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiImport).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiExport).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstImages
        mvMenuItems(WebMenuItems.wmiNewImage).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstPages
        mvMenuItems(WebMenuItems.wmiNewPage).SetContextItemVisible(Me, True)
        'mvMenuItems(WebMenuItems.wmiNewPageItem).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstPage
        mvMenuItems(WebMenuItems.wmiNewPage).SetContextItemVisible(Me, True)
        If mvID > 0 Then
          mvMenuItems(WebMenuItems.wmiNewPageItem).SetContextItemVisible(Me, True)
          mvMenuItems(WebMenuItems.wmiCopyPage).SetContextItemVisible(Me, True)
          mvMenuItems(WebMenuItems.wmiDeletePage).SetContextItemVisible(Me, True)
        End If
      Case CareWebAccess.XMLWebDataSelectionTypes.wstPageItem
        mvMenuItems(WebMenuItems.wmiNewPageItem).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiDeletePageItem).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiMoveItemUp).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiMoveItemDown).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstMenus
        mvMenuItems(WebMenuItems.wmiNewMenu).SetContextItemVisible(Me, True)
        'mvMenuItems(WebMenuItems.wmiNewMenuItem).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstMenu
        mvMenuItems(WebMenuItems.wmiNewMenu).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiNewMenuItem).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiDeleteMenu).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstMenuItem
        mvMenuItems(WebMenuItems.wmiNewMenuItem).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiDeleteMenuItem).SetContextItemVisible(Me, True)
        mvMenuItems(WebMenuItems.wmiAddPage).SetContextItemVisible(Me, True)
      Case CareWebAccess.XMLWebDataSelectionTypes.wstDocuments
        mvMenuItems(WebMenuItems.wmiNewDocument).SetContextItemVisible(Me, True)
      Case Else
        e.Cancel = True
    End Select
  End Sub
End Class
