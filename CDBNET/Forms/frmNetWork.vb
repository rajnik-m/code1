Public Class frmNetWork

  Public Sub New(ByVal pContactInfo As ContactInfo, ByVal pType As CareServices.XMLContactDataSelectionTypes)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pContactInfo, pType)
  End Sub

  Private mvContactInfo As ContactInfo
  Private mvType As CareServices.XMLContactDataSelectionTypes
  Private mvBrowserMenu As BrowserMenu

  Private Sub InitialiseControls(ByVal pContactInfo As ContactInfo, ByVal pType As CareServices.XMLContactDataSelectionTypes)
    SetControlTheme()
    MainHelper.SetMDIParent(Me)
    Select Case pType
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom
        Me.Text = String.Format(ControlText.frmNetworkFrom, pContactInfo.ContactName)
      Case CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo
        Me.Text = String.Format(ControlText.frmNetworkTo, pContactInfo.ContactName)
    End Select
    tvw.ImageList = MainHelper.ImageProvider.NewTreeViewImages
    mvContactInfo = pContactInfo
    mvType = pType
    Dim vNewNode As TreeNode = New TreeNode(mvContactInfo.ContactName)
    If DataHelper.ContactAndOrganisationGroups.ContainsKey(mvContactInfo.ContactGroup) Then
      vNewNode.ImageIndex = DataHelper.ContactAndOrganisationGroups(mvContactInfo.ContactGroup).ImageIndex
    Else
      vNewNode.ImageIndex = 0
    End If
    vNewNode.SelectedImageIndex = vNewNode.ImageIndex
    vNewNode.Tag = mvContactInfo.ContactNumber
    tvw.Nodes.Add(vNewNode)
    vNewNode.Nodes.Add("_DUMMY_")
    mvBrowserMenu = New BrowserMenu(Nothing)
    tvw.ContextMenuStrip = mvBrowserMenu
  End Sub

  Private Sub GetLinks(ByVal pNode As TreeNode, ByVal pContactNumber As Integer)
    Dim vContactInfo As New ContactInfo(pContactNumber)
    If FormHelper.CheckAccessRights(vContactInfo) Then
      Dim vDataset As DataSet = DataHelper.GetContactData(mvType, pContactNumber)
      If vDataset.Tables.Contains("DataRow") Then
        For Each vRow As DataRow In vDataset.Tables("DataRow").Rows
          Dim vHistoric As Boolean = vRow("Historical").ToString.Length > 0
          If Not (Settings.HideHistoricNetwork AndAlso vHistoric) Then
            Dim vText As String = String.Format("{0} ({1})", vRow.Item("ContactName").ToString, vRow.Item("RelationshipDesc").ToString)
            Dim vNewNode As TreeNode = New TreeNode(vText)
            Dim vGroup As String = vRow.Item("ContactGroup").ToString
            If DataHelper.ContactAndOrganisationGroups.ContainsKey(vGroup) Then
              vNewNode.ImageIndex = DataHelper.ContactAndOrganisationGroups(vGroup).ImageIndex
            Else
              vNewNode.ImageIndex = 0
            End If
            vNewNode.SelectedImageIndex = vNewNode.ImageIndex
            vNewNode.Tag = vRow.Item("ContactNumber")
            If vHistoric Then vNewNode.ForeColor = Color.Gray
            pNode.Nodes.Add(vNewNode)
            vNewNode.Nodes.Add("_DUMMY_")
          End If
        Next
      End If
    End If
  End Sub

  Private Sub tvw_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles tvw.BeforeExpand
    Dim vNode As TreeNode = e.Node
    If Not vNode Is Nothing Then
      Dim vFound As Boolean = CheckNodes(tvw.Nodes(0), vNode)
      If vNode.Nodes.Count = 1 And vNode.Nodes(0).Text = "_DUMMY_" Then
        vNode.Nodes.RemoveAt(0)
        If Not vFound Then
          GetLinks(vNode, CInt(vNode.Tag))
        End If
      End If
    End If
  End Sub

  Private Sub tvw_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvw.MouseDown
    If e.Button = System.Windows.Forms.MouseButtons.Right Then
      Dim vNode As TreeNode = tvw.GetNodeAt(e.X, e.Y)
      tvw.SelectedNode = vNode
      If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
        Dim vContactNumber As Integer = CInt(vNode.Tag.ToString)
        If vContactNumber > 0 Then
          Dim vContactInfo As New ContactInfo(vContactNumber)
          mvBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
          mvBrowserMenu.ItemNumber = vContactNumber
          mvBrowserMenu.ItemDescription = vContactInfo.ContactName
          mvBrowserMenu.GroupCode = vContactInfo.ContactGroup
        End If
      End If
    End If
  End Sub

  Private Function CheckNodes(ByVal pCheckNode As TreeNode, ByVal pExpandNode As TreeNode) As Boolean
    Dim vFound As Boolean
    For Each vCheckNode As TreeNode In pCheckNode.Nodes
      If Not (vCheckNode Is pExpandNode) AndAlso CInt(vCheckNode.Tag) = CInt(pExpandNode.Tag) Then
        If ShowQuestion(QuestionMessages.QmNetworkContactExists, MessageBoxButtons.YesNo, vCheckNode.Text) = System.Windows.Forms.DialogResult.Yes Then
          tvw.SelectedNode = vCheckNode
        End If
        vFound = True
      Else
        If vCheckNode.Nodes.Count > 0 Then vFound = CheckNodes(vCheckNode, pExpandNode)
      End If
      If vFound Then Return vFound
    Next
  End Function

  Public ReadOnly Property SelectedContactNumber() As Integer
    Get
      Return mvBrowserMenu.ItemNumber
    End Get
  End Property

  Public Overloads Sub Refresh()
    For vIndex As Integer = 0 To tvw.Nodes.Count - 1
      Me.tvw.Nodes.Remove(tvw.Nodes.Item(vIndex))
    Next
    InitialiseControls(mvContactInfo, mvType)
  End Sub

End Class