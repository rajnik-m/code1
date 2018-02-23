Friend Class FastDataEntryMenu
  Inherits ContextMenuStrip

  Friend Enum FDEBrowserMenuItems
    AddModule
    EditPage
    'The items below are sub-menu items
    ActivityDisplay
    AddDonationCC
    AddMemberDD
    AddTransaction
    AddressDisplay
    CommunicationsDisplay
    ContactSelection
    DisplayLabel
    GiftAidDisplay
    SuppressionDisplay
    AddRegularDonation
    Telemarketing
    ProductSale
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)
  Private mvModuleMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Event MenuSelected(ByVal pMenuItem As ToolStripMenuItem, ByVal pItem As FDEBrowserMenuItems)

  Public Sub New()
    MyBase.New()

    With mvMenuItems
      .Add(FDEBrowserMenuItems.AddModule.ToString, New MenuToolbarCommand(FDEBrowserMenuItems.AddModule.ToString, ControlText.MnuFDEFrmAddModule, FDEBrowserMenuItems.AddModule, "SCFPAM"))
      .Add(FDEBrowserMenuItems.EditPage.ToString, New MenuToolbarCommand(FDEBrowserMenuItems.EditPage.ToString, ControlText.MnuFDEFrmEditPage, FDEBrowserMenuItems.EditPage, "SCFPEP"))
    End With

    Dim vDT As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtFastDataEntryUserControls)
    If vDT IsNot Nothing Then
      For Each vRow As DataRow In vDT.Rows
        With mvModuleMenuItems
          Select Case vRow.Item("FdeUserControl").ToString
            Case "ACTIVITYDISPLAY"
              .Add(FDEBrowserMenuItems.ActivityDisplay.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.ActivityDisplay, "SCFPDA"))
            Case "ADDDONATIONCC"
              .Add(FDEBrowserMenuItems.AddDonationCC.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.AddDonationCC, "SCFPAD"))
            Case "ADDMEMBERDD"
              .Add(FDEBrowserMenuItems.AddMemberDD.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.AddMemberDD, "SCFPAM"))
            Case "ADDREGULARDON"
              .Add(FDEBrowserMenuItems.AddRegularDonation.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.AddRegularDonation, "SCFPAR"))
            Case "ADDTRANSACTIONDETAILS"
              .Add(FDEBrowserMenuItems.AddTransaction.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.AddTransaction, "SCFPTR"))
            Case "ADDRESSDISPLAY"
              .Add(FDEBrowserMenuItems.AddressDisplay.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.AddressDisplay, "SCFPDR"))
            Case "COMMUNICATIONSDISPLAY"
              .Add(FDEBrowserMenuItems.CommunicationsDisplay.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.CommunicationsDisplay, "SCFPCM"))
            Case "CONTACTSELECTION"
              .Add(FDEBrowserMenuItems.ContactSelection.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.ContactSelection, "SCFPCS"))
            Case "DISPLAYLABEL"
              .Add(FDEBrowserMenuItems.DisplayLabel.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.DisplayLabel, "SCFPDL"))
            Case "GIFTAIDDISPLAY"
              .Add(FDEBrowserMenuItems.GiftAidDisplay.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.GiftAidDisplay, "SCFPGA"))
            Case "SUPPRESSIONDISPLAY"
              .Add(FDEBrowserMenuItems.SuppressionDisplay.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.SuppressionDisplay, "SCFPSD"))
            Case "TELEMARKETING"
              .Add(FDEBrowserMenuItems.Telemarketing.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.Telemarketing, "SCFPTM"))
            Case "PRODUCTSALE"
              .Add(FDEBrowserMenuItems.ProductSale.ToString, New MenuToolbarCommand(vRow.Item("FdeUserControl").ToString, vRow.Item("ControlTitle").ToString, FDEBrowserMenuItems.ProductSale, "SCFPPS"))
          End Select
        End With
      Next
    End If

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next

    Dim vAddModuleItem As ToolStripMenuItem = DirectCast(Me.Items(FDEBrowserMenuItems.AddModule), ToolStripMenuItem)
    With vAddModuleItem.DropDownItems
      For Each vItem As MenuToolbarCommand In mvModuleMenuItems
        vItem.OnClick = AddressOf MenuHandler
        .Add(vItem.MenuStripItem)
      Next
    End With
    
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
    MenuToolbarCommand.SetAccessControl(mvModuleMenuItems)

  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      Dim vMenu As MenuToolbarCommand = DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand)
      Dim vFDEItem As FDEBrowserMenuItems = CType(vMenu.CommandID, FDEBrowserMenuItems)
      RaiseEvent MenuSelected(DirectCast(sender, ToolStripMenuItem), vFDEItem)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

End Class
