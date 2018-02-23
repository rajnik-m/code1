Public Class frmAccessControl
  Private mvAccessGroupsTable As New DataTable
  Private mvAccessAreasTable As New DataTable
  Private mvAccessControlTable As New DataTable
  Private WithEvents mvAccessControlMenu As AccessControlMenu
  Public Sub New(ByVal pAccessControlVersion As Integer)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pAccessControlVersion)
  End Sub

  Private Sub InitialiseControls(ByVal pAccessControlVersion As Integer)
    Try
      mvAccessControlMenu = New AccessControlMenu()
      CheckAccessControl(pAccessControlVersion)
      tvw.ImageList = AppHelper.ImageProvider.NewTreeViewImages
      GetAccessControlTree()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try

  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      Me.Close()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub CheckAccessControl(ByVal pAccessControlVersion As Integer)
    Dim vNewVersion As Integer
    Dim vParams As New ParameterList(True)
    vNewVersion = pAccessControlVersion
    vParams("ConfigName") = "access_control_version"
    vParams("NewVersion") = vNewVersion.ToString
    DataHelper.CreateAccessControlData(vParams)
  End Sub

  ''' <summary>
  ''' This function loads the AccessControlGroups in tree
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetAccessControlTree()
    Dim vAreaNode As New TreeNode
    Dim vMenuNode As New TreeNode
    Dim vItemNode As New TreeNode
    Dim vNodeInfo As ControlNodeInfo
    Dim vAccessItemTable As New DataTable
    Dim vGroups As New ParameterList
    Dim vParentName As String
    Dim vKey As String

    tvw.Nodes.Clear()
    'Get all the access_control groups
    mvAccessGroupsTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtAccessControlGroups)
    If mvAccessGroupsTable IsNot Nothing Then
      With mvAccessGroupsTable
        For vGroupRowCounter As Integer = 0 To mvAccessGroupsTable.Rows.Count - 1
          Dim vNode As TreeNode = New TreeNode()
          vKey = mvAccessGroupsTable.Rows(vGroupRowCounter).Item("AccessControlGroup").ToString()
          vParentName = mvAccessGroupsTable.Rows(vGroupRowCounter).Item("AccessControlGroupDesc").ToString()
          vNodeInfo = New ControlNodeInfo(ControlNodeType.Group, vKey)
          'Group Nodes
          vNode = tvw.Nodes.Add(vKey, vParentName)
          vNode.Tag = vNodeInfo
          'Get all the access_control areas
          mvAccessAreasTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtAccessControlAreas)
          Dim vParameterList As New ParameterList(True)
          vParameterList("AccessControlGroup") = vNodeInfo.ControlGroup
          mvAccessControlTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtAccessControl, vParameterList)
          BuildAreaNodes(vNode)
        Next
      End With
    End If
  End Sub

  Private Sub tvw_BeforeExpand(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles tvw.BeforeExpand
    Try
      Dim vNodeInfo As ControlNodeInfo = CType(e.Node.Tag, ControlNodeInfo)
      If Not e.Node.Nodes.Item("_DUMMY") Is Nothing Then
        e.Node.Nodes.RemoveByKey("_DUMMY")
        Select Case vNodeInfo.NodeType
          Case ControlNodeType.Area
            BuildMenuNodes(e.Node)
          Case ControlNodeType.Menu
            BuildSubMenuNodes(e.Node)
          Case ControlNodeType.SubMenu
            BuildItemNodes(e.Node)
        End Select
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  ''' <summary>
  ''' This function loads the Area Nodes
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <remarks></remarks>
  Private Sub BuildAreaNodes(ByVal pNode As TreeNode)
    Dim vKey As String
    Dim vNodeText As String
    Dim vNodeInfo As ControlNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    For VRowCounter As Integer = 0 To mvAccessAreasTable.Rows.Count - 1
      vKey = mvAccessAreasTable.Rows(VRowCounter).Item("AccessControlArea").ToString()
      vNodeText = mvAccessAreasTable.Rows(VRowCounter).Item("AccessControlAreaDesc").ToString()
      If vKey.Length = 2 Then
        Dim vSubNodeInfo As ControlNodeInfo
        vSubNodeInfo = New ControlNodeInfo(ControlNodeType.Area, vNodeInfo.ControlGroup, vKey)
        pNode.Nodes.Add(vKey, vNodeText)
        pNode.Nodes(vKey).Tag = vSubNodeInfo
        BuildMenuNodes(pNode.Nodes.Item(vKey))
      End If
    Next
  End Sub
  ''' <summary>
  ''' This function loads the Menu Nodes
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <remarks></remarks>
  Private Sub BuildMenuNodes(ByVal pNode As TreeNode)
    Dim vKey As String
    Dim vNodeText As String
    Dim vNodeInfo As ControlNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    For VRowCounter As Integer = 0 To mvAccessAreasTable.Rows.Count - 1
      vKey = mvAccessAreasTable.Rows(VRowCounter).Item("AccessControlArea").ToString()
      vNodeText = mvAccessAreasTable.Rows(VRowCounter).Item("AccessControlAreaDesc").ToString()
      If vKey.ToString.Length = 4 Then
        Dim vMenuNodeInfo As ControlNodeInfo
        vMenuNodeInfo = New ControlNodeInfo(ControlNodeType.Menu, vNodeInfo.ControlGroup, vNodeInfo.ControlArea, vKey)
        pNode.Nodes.Add(vKey, vNodeText)
        pNode.Nodes(vKey).Tag = vMenuNodeInfo
        BuildSubMenuNodes(pNode.Nodes.Item(vKey))
      End If
    Next
    'Adding an extra menu Document Popup which is not in Access_Control_Area
    Dim vMenuNodeInfo1 As ControlNodeInfo
    vMenuNodeInfo1 = New ControlNodeInfo(ControlNodeType.Menu, vNodeInfo.ControlGroup, vNodeInfo.ControlArea, "CDDP")
    pNode.Nodes.Add("CDDP", "Document Popup")
    pNode.Nodes("CDDP").Tag = vMenuNodeInfo1
    BuildSubMenuNodes(pNode.Nodes.Item("CDDP"))
  End Sub
  ''' <summary>
  ''' This function loads Sub-Menu Nodes
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <remarks></remarks>
  Private Sub BuildSubMenuNodes(ByVal pNode As TreeNode)
    Dim vKey As String
    Dim vNodeText As String
    Dim vNodeInfo As ControlNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    For vRowCounter As Integer = 0 To mvAccessAreasTable.Rows.Count - 1
      pNode.Nodes.RemoveByKey("_DUMMY")
      vKey = mvAccessAreasTable.Rows(vRowCounter).Item("AccessControlArea").ToString()
      vNodeText = mvAccessAreasTable.Rows(vRowCounter).Item("AccessControlAreaDesc").ToString()
      If vKey.ToString.Length = 6 Then
        If vKey.ToString.StartsWith(pNode.Name) Then
          Dim vSubMenuNodeInfo As ControlNodeInfo
          vSubMenuNodeInfo = New ControlNodeInfo(ControlNodeType.SubMenu, vNodeInfo.ControlGroup, vNodeInfo.ControlArea, vNodeInfo.ControlMenu, vKey)
          pNode.Nodes.Add(vKey, vNodeText)
          pNode.Nodes(vKey).Tag = vSubMenuNodeInfo
          BuildItemNodes(pNode.Nodes.Item(vKey))
        End If
      End If
    Next
    BuildItemNodes(pNode)

  End Sub
  ''' <summary>
  ''' This function loads Item Nodes
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <remarks></remarks>
  Private Sub BuildItemNodes(ByVal pNode As TreeNode)
    Dim vKey As String
    Dim vNodeText As String
    Dim vAccessLevel As String
    Dim vDefaultAccessLevel As String
    Dim vNodeInfo As ControlNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    For vRowCounter As Integer = 0 To mvAccessControlTable.Rows.Count - 1
      pNode.Nodes.RemoveByKey("_DUMMY")
      vKey = mvAccessControlTable.Rows(vRowCounter).Item("AccessControlItem").ToString()
      vNodeText = mvAccessControlTable.Rows(vRowCounter).Item("AccessControlItemDesc").ToString()
      Dim vAccessControlArea As String = mvAccessControlTable.Rows(vRowCounter).Item("AccessControlArea").ToString()
      vDefaultAccessLevel = mvAccessControlTable.Rows(vRowCounter).Item("AccessLevel").ToString()
      If mvAccessControlTable.Rows(vRowCounter).Item("NewAccessLevel").ToString().Length = 0 Then
        vAccessLevel = mvAccessControlTable.Rows(vRowCounter).Item("AccessLevel").ToString()
      Else
        vAccessLevel = mvAccessControlTable.Rows(vRowCounter).Item("NewAccessLevel").ToString()
      End If
      'Below or condition is written because "CDFP", "CDAP" & "CDEV" are common for rich client and smart client
      If vAccessControlArea = pNode.Name Or (pNode.Name = "SCFP" And ((vAccessControlArea = "CDFP" And vKey <> "CDFPPR" And vKey <> "CDFAPO") Or vAccessControlArea = "CDAP" _
      Or (vAccessControlArea = "CDEV" And (vKey = "CDEVFL" Or vKey = "CDEVCA" Or vKey = "CDEVSI" Or vKey = "CDEVWL")))) Then
        Dim vItemNodeInfo As ControlNodeInfo
        vItemNodeInfo = New ControlNodeInfo(ControlNodeType.Item, vNodeInfo.ControlGroup, vNodeInfo.ControlArea, vNodeInfo.ControlMenu, vNodeInfo.ControlSubMenu, vKey, vAccessLevel, vDefaultAccessLevel)
        pNode.Nodes.Add(vKey, vNodeText)

        'Setting Icon
        pNode.Nodes.Item(vKey).ImageKey = GetSelectedImageKey(vAccessLevel)
        pNode.Nodes.Item(vKey).SelectedImageKey = pNode.Nodes.Item(vKey).ImageKey
        pNode.Nodes(vKey).Tag = vItemNodeInfo
        BuildItemNodes(pNode.Nodes.Item(vKey))
      End If
    Next
  End Sub

  Private Function GetSelectedImageKey(ByVal pAccessLevel As String) As String
    Return "AccessControl" & pAccessLevel
  End Function

  Private Sub tvw_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvw.MouseDown
    Try
      If e.Button = System.Windows.Forms.MouseButtons.Right Then
        Dim vNode As TreeNode = tvw.GetNodeAt(e.X, e.Y)
        tvw.SelectedNode = vNode
        If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
          tvw.ContextMenuStrip = mvAccessControlMenu
          If tvw.ContextMenuStrip IsNot Nothing Then
            Dim vNodeInfo As ControlNodeInfo = CType(vNode.Tag, ControlNodeInfo)
            mvAccessControlMenu.ParentNode = vNode
            mvAccessControlMenu.ControlNodeInfo = vNodeInfo
          End If
        Else
          tvw.ContextMenuStrip = Nothing
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  Private Sub ShowNodeAccessLevel(ByVal pNode As TreeNode)
    Dim vNodeInfo As ControlNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    Select Case vNodeInfo.NodeType
      Case ControlNodeType.Group, ControlNodeType.Area, ControlNodeType.Menu, ControlNodeType.SubMenu
        txtAccessLevelReqd.Text = String.Empty
      Case ControlNodeType.Item
        Select Case vNodeInfo.AccessLevel
          Case "D"
            txtAccessLevelReqd.Text = ControlText.TxtDBAdmin
          Case "U"
            txtAccessLevelReqd.Text = ControlText.TxtUser
          Case "S"
            txtAccessLevelReqd.Text = ControlText.TxtSupervisor
          Case "R"
            txtAccessLevelReqd.Text = ControlText.TxtReadOnly
          Case "N"
            txtAccessLevelReqd.Text = ControlText.TxtNotAvailableToUser
        End Select
        If vNodeInfo.AccessLevel = vNodeInfo.DefaultAccessLevel Then
          txtAccessLevelReqd.Text = txtAccessLevelReqd.Text & " (Default Setting)"
        End If

    End Select
  End Sub
  Private Sub tvw_NodeMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvw.NodeMouseClick
    Try
      ShowNodeAccessLevel(e.Node)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub mvAccessControlMenu_MenuActionCompleted(ByVal pItem As AccessControlMenu.AccessControlMenuItems) Handles mvAccessControlMenu.MenuActionCompleted
    Try
      Dim vNodeInfo As ControlNodeInfo
      vNodeInfo = CType(tvw.SelectedNode.Tag, ControlNodeInfo)
      Select Case pItem
        Case AccessControlMenu.AccessControlMenuItems.acmiDeleteGroup, AccessControlMenu.AccessControlMenuItems.acmiNewGroup
          GetAccessControlTree()
        Case AccessControlMenu.AccessControlMenuItems.acmiAdministrator
          UpdateAccessControlImage(tvw.SelectedNode, "D")
        Case AccessControlMenu.AccessControlMenuItems.acmiSupervisor
          UpdateAccessControlImage(tvw.SelectedNode, "S")
        Case AccessControlMenu.AccessControlMenuItems.acmiUser
          UpdateAccessControlImage(tvw.SelectedNode, "U")
        Case AccessControlMenu.AccessControlMenuItems.acmiReadOnly
          UpdateAccessControlImage(tvw.SelectedNode, "R")
        Case AccessControlMenu.AccessControlMenuItems.acmiNone
          UpdateAccessControlImage(tvw.SelectedNode, "N")
        Case AccessControlMenu.AccessControlMenuItems.acmiDefault
          UpdateAccessControlImage(tvw.SelectedNode, vNodeInfo.DefaultAccessLevel)
          SetDefault(tvw.SelectedNode)
      End Select
      If vNodeInfo.NodeType = ControlNodeType.Item Then
        ShowNodeAccessLevel(tvw.SelectedNode)
      End If

    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  Private Sub UpdateAccessControlImage(ByVal pNode As TreeNode, ByVal pAccessLevel As String)
    Dim vNodeInfo As ControlNodeInfo
    Dim vParam As New ParameterList(True)
    vNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    If vNodeInfo.NodeType = ControlNodeType.Item Then
      vNodeInfo.AccessLevel = pAccessLevel
      pNode.Tag = vNodeInfo
      pNode.ImageKey = GetSelectedImageKey(pAccessLevel)
      pNode.SelectedImageKey = pNode.ImageKey
    Else
      For vCtr As Integer = 0 To pNode.Nodes.Count - 1
        UpdateAccessControlImage(pNode.Nodes(vCtr), pAccessLevel)
      Next
    End If
  End Sub

  Private Sub SetDefault(ByVal pNode As TreeNode)
    Dim vNodeInfo As ControlNodeInfo
    vNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    Select Case vNodeInfo.NodeType
      Case ControlNodeType.Group
        For Each vNode As TreeNode In pNode.Nodes
          SetDefault(vNode)
        Next
      Case ControlNodeType.Area
        For Each vNode As TreeNode In pNode.Nodes
          SetDefault(vNode)
        Next
      Case ControlNodeType.Menu
        For Each vNode As TreeNode In pNode.Nodes
          SetDefault(vNode)
        Next
      Case ControlNodeType.SubMenu
        For Each vNode As TreeNode In pNode.Nodes
          SetDefault(vNode)
        Next
      Case ControlNodeType.Item
        UpdateAccessControlImage(pNode, vNodeInfo.DefaultAccessLevel)
    End Select
  End Sub
End Class



Public Enum ControlNodeType
  Group
  Area
  Menu
  SubMenu
  Item
  AccessLevel
End Enum
Public Class ControlNodeInfo
  Private mvNodeType As ControlNodeType
  Private mvControlGroup As String = ""
  Private mvControlArea As String = ""
  Private mvControlMenu As String = ""
  Private mvControlSubMenu As String = ""
  Private mvControlItem As String = ""
  Private mvAccessLevel As String = ""
  Private mvDefaultAccessLevel As String = ""

  Public Sub New(ByVal pNodeType As ControlNodeType, Optional ByVal pControlGroup As String = "", Optional ByVal pControlArea As String = "", Optional ByVal pControlMenu As String = "", Optional ByVal pControlSubMenu As String = "", Optional ByVal pControlItem As String = "", Optional ByVal pAccessLevel As String = "", Optional ByVal pDefaultAccessLevel As String = "")
    mvNodeType = pNodeType
    mvControlGroup = pControlGroup
    mvControlArea = pControlArea
    mvControlMenu = pControlMenu
    mvControlSubMenu = pControlSubMenu
    mvControlItem = pControlItem
    mvAccessLevel = pAccessLevel
    mvDefaultAccessLevel = pDefaultAccessLevel
  End Sub
  Public ReadOnly Property NodeType() As Integer
    Get
      Return mvNodeType
    End Get
  End Property
  Public ReadOnly Property ControlGroup() As String
    Get
      Return mvControlGroup
    End Get
  End Property
  Public ReadOnly Property ControlArea() As String
    Get
      Return mvControlArea
    End Get
  End Property
  Public ReadOnly Property ControlMenu() As String
    Get
      Return mvControlMenu
    End Get
  End Property
  Public ReadOnly Property ControlSubMenu() As String
    Get
      Return mvControlSubMenu
    End Get
  End Property
  Public ReadOnly Property ControlItem() As String
    Get
      Return mvControlItem
    End Get
  End Property
  Public Property AccessLevel() As String
    Get
      Return mvAccessLevel
    End Get
    Set(ByVal Value As String)
      mvAccessLevel = Value
    End Set
  End Property
  Public ReadOnly Property DefaultAccessLevel() As String
    Get
      Return mvDefaultAccessLevel
    End Get
  End Property
End Class