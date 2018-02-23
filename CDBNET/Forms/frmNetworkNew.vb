Public Class frmNetworkNew
  Inherits MaintenanceParentForm

  Private WithEvents mvNetworkTreeview As NetworkTreeView
  Private mvContactInfo As ContactInfo
  Private mvType As CareServices.XMLContactDataSelectionTypes

  Public Sub New()
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
  End Sub

  Public Sub New(ByVal pContactInfo As ContactInfo, ByVal pType As CareServices.XMLContactDataSelectionTypes)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    mvContactInfo = pContactInfo
    mvType = pType

    Dim vPrefix As String = ""
    Dim vEntityGroup As EntityGroup
    If DataHelper.ContactAndOrganisationGroups.ContainsKey(pContactInfo.ContactGroup) Then
      vEntityGroup = DataHelper.ContactAndOrganisationGroups(pContactInfo.ContactGroup)
      If vEntityGroup.ImageIndex < MainHelper.ImageProvider.NewTreeViewImages.Images.Count Then
        Me.Icon = Drawing.Icon.FromHandle(CType(MainHelper.ImageProvider.NewTreeViewImages.Images(vEntityGroup.ImageIndex), Bitmap).GetHicon)
        vPrefix = vEntityGroup.GroupName & ": "
      Else
        Me.Icon = vEntityGroup.Icon
      End If
    End If
    Me.Text = vPrefix & pContactInfo.ContactName

    mvNetworkTreeview = New NetworkTreeView(pContactInfo, pType)
    mvNetworkTreeview.Visible = True
    mvNetworkTreeview.Dock = DockStyle.Fill
    mvNetworkTreeview.SetBrowserMenu()
    pnlTreeView.Controls.Add(mvNetworkTreeview)
  End Sub

  Public Sub InitialiseControls(ByVal pContactInfo As ContactInfo, ByVal pType As CareServices.XMLContactDataSelectionTypes)
    mvContactInfo = pContactInfo
    mvType = pType
  End Sub

  Public Overloads Sub Refresh()
    mvNetworkTreeview.Refresh()
  End Sub

  Private Sub mvNetworkTreeView_SetBrowserMenuContext(ByVal psender As Object) Handles mvNetworkTreeview.SetBrowserMenuContext
    MainHelper.SetBrowserMenu(psender, Me)
  End Sub

  Private Sub mvNetworkTreeView_SetBrowserProperty(ByVal psender As Object, ByVal pNodeInfo As CDBNETCL.TreeViewNodeInfo) Handles mvNetworkTreeview.SetBrowserProperty
    Try
      If CType(psender, TreeView) IsNot Nothing OrElse CType(psender, VistaTreeView) IsNot Nothing Then
        Dim vBrowserMenu As BrowserMenu = CType(CType(psender, VistaTreeView).ContextMenuStrip, BrowserMenu)
        If vBrowserMenu IsNot Nothing Then
          If Not pNodeInfo.IsFolder AndAlso pNodeInfo.AddressNumber = 0 Then
            vBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
            vBrowserMenu.ItemNumber = pNodeInfo.ContactNumber
            vBrowserMenu.ItemDescription = pNodeInfo.ContactName
            vBrowserMenu.GroupCode = pNodeInfo.Group
          Else
            vBrowserMenu.EntityType = HistoryEntityTypes.hetNone
            vBrowserMenu.GroupCode = String.Empty
          End If
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
End Class