Public Class AccessControlMenu
  Inherits ContextMenuStrip

  Private mvParent As MaintenanceParentForm
  Private mvNodeInfo As ControlNodeInfo
  Private mvParentNode As TreeNode
  Public Event MenuActionCompleted(ByVal pItem As AccessControlMenuItems)

  Public Enum AccessControlMenuItems
    acmiAdministrator
    acmiSupervisor
    acmiUser
    acmiReadOnly
    acmiNone
    acmiDefault
    acmiNewGroup
    acmiDeleteGroup
  End Enum
  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New()
    MyBase.New()
    With mvMenuItems
      .Add(AccessControlMenuItems.acmiAdministrator.ToString, New MenuToolbarCommand("Administrator", ControlText.MnuAdministrator, AccessControlMenuItems.acmiAdministrator))
      .Add(AccessControlMenuItems.acmiSupervisor.ToString, New MenuToolbarCommand("Supervisor", ControlText.MnuSupervisor, AccessControlMenuItems.acmiSupervisor))
      .Add(AccessControlMenuItems.acmiUser.ToString, New MenuToolbarCommand("User", ControlText.MnuUser, AccessControlMenuItems.acmiUser))
      .Add(AccessControlMenuItems.acmiReadOnly.ToString, New MenuToolbarCommand("ReadOnly", ControlText.MnuReadOnly, AccessControlMenuItems.acmiReadOnly))
      .Add(AccessControlMenuItems.acmiNone.ToString, New MenuToolbarCommand("None", ControlText.MnuNone, AccessControlMenuItems.acmiNone))
      .Add(AccessControlMenuItems.acmiDefault.ToString, New MenuToolbarCommand("Default", ControlText.MnuDefault, AccessControlMenuItems.acmiDefault))
      .Add(AccessControlMenuItems.acmiNewGroup.ToString, New MenuToolbarCommand("NewGroup", ControlText.MnuNewGroup, AccessControlMenuItems.acmiNewGroup))
      .Add(AccessControlMenuItems.acmiDeleteGroup.ToString, New MenuToolbarCommand("DeleteGroup", ControlText.MnuDeleteGroup, AccessControlMenuItems.acmiDeleteGroup))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
  End Sub

  Public Property ParentNode() As TreeNode
    Get
      Return mvParentNode
    End Get
    Set(ByVal Value As TreeNode)
      mvParentNode = Value
    End Set
  End Property

  Public Property ControlNodeInfo() As ControlNodeInfo
    Get
      Return mvNodeInfo
    End Get
    Set(ByVal Value As ControlNodeInfo)
      mvNodeInfo = Value
    End Set
  End Property

  Private Sub UpdateAccessControlItem(ByVal pNode As TreeNode, ByVal pAccessLevel As String)
    Dim vNodeInfo As ControlNodeInfo
    Dim vParam As New ParameterList(True)
    vNodeInfo = CType(pNode.Tag, ControlNodeInfo)
    If vNodeInfo.NodeType = ControlNodeType.Item Then
      If pAccessLevel = "" Then
        vParam("AccessLevel") = vNodeInfo.DefaultAccessLevel
      Else
        vParam("AccessLevel") = pAccessLevel
      End If
      vParam("AccessControlGroup") = vNodeInfo.ControlGroup
      vParam("AccessControlItem") = vNodeInfo.ControlItem

      DataHelper.UpdateAccessControlItem(vParam)
    Else
      For vCtr As Integer = 0 To pNode.Nodes.Count - 1
        UpdateAccessControlItem(pNode.Nodes(vCtr), pAccessLevel)
      Next
    End If
  End Sub
  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As AccessControlMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, AccessControlMenuItems)
      Dim vCurrentTableName As String = String.Empty
      Dim vParams As New ParameterList(True)
      Dim vCriteriaList As New ParameterList

      Select Case vMenuItem
        Case AccessControlMenuItems.acmiDeleteGroup
          If ShowQuestion(QuestionMessages.QmDeleteAccessControlGroup, MessageBoxButtons.YesNo) = DialogResult.Yes Then
            vParams("AccessControlGroup") = mvNodeInfo.ControlGroup
            DataHelper.DeleteAccessControlGroup(vParams)
            RaiseEvent MenuActionCompleted(vMenuItem)
          End If
        Case AccessControlMenuItems.acmiAdministrator
          UpdateAccessControlItem(mvParentNode, "D")
          RaiseEvent MenuActionCompleted(vMenuItem)
        Case AccessControlMenuItems.acmiSupervisor
          UpdateAccessControlItem(mvParentNode, "S")
          RaiseEvent MenuActionCompleted(vMenuItem)
        Case AccessControlMenuItems.acmiUser
          UpdateAccessControlItem(mvParentNode, "U")
          RaiseEvent MenuActionCompleted(vMenuItem)
        Case AccessControlMenuItems.acmiReadOnly
          UpdateAccessControlItem(mvParentNode, "R")
          RaiseEvent MenuActionCompleted(vMenuItem)
        Case AccessControlMenuItems.acmiNone
          UpdateAccessControlItem(mvParentNode, "N")
          RaiseEvent MenuActionCompleted(vMenuItem)
        Case AccessControlMenuItems.acmiDefault
          UpdateAccessControlItem(mvParentNode, mvNodeInfo.DefaultAccessLevel)

          RaiseEvent MenuActionCompleted(vMenuItem)
        Case AccessControlMenuItems.acmiNewGroup
          vCurrentTableName = "access_control_groups"
          vParams("MaintenanceTableName") = vCurrentTableName
          Dim vForm As New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vCurrentTableName, vParams, Nothing)
          vForm.Text = ControlText.FrmAddToAccessControlGroups
          If vForm.ShowDialog() = DialogResult.OK Then
            RaiseEvent MenuActionCompleted(vMenuItem)
          End If
      End Select

    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub AccessControlMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    Dim vCursor As New BusyCursor
    Try
      Select Case mvNodeInfo.NodeType
        Case ControlNodeType.Group
          Me.Items(AccessControlMenuItems.acmiDefault).Visible = True
          Me.Items(AccessControlMenuItems.acmiNewGroup).Visible = True
          Me.Items(AccessControlMenuItems.acmiDeleteGroup).Visible = True
          Me.Items(AccessControlMenuItems.acmiAdministrator).Visible = False
          Me.Items(AccessControlMenuItems.acmiNone).Visible = False
          Me.Items(AccessControlMenuItems.acmiReadOnly).Visible = False
          Me.Items(AccessControlMenuItems.acmiSupervisor).Visible = False
          Me.Items(AccessControlMenuItems.acmiUser).Visible = False
          If Not mvNodeInfo.ControlGroup = "MAIN" Then
            Me.Items(AccessControlMenuItems.acmiDeleteGroup).Enabled = True
          Else
            Me.Items(AccessControlMenuItems.acmiDeleteGroup).Enabled = False
          End If
        Case ControlNodeType.Area, ControlNodeType.Menu, ControlNodeType.SubMenu, ControlNodeType.Item
          Me.Items(AccessControlMenuItems.acmiAdministrator).Visible = True
          Me.Items(AccessControlMenuItems.acmiDefault).Visible = True
          Me.Items(AccessControlMenuItems.acmiNewGroup).Visible = False
          Me.Items(AccessControlMenuItems.acmiDeleteGroup).Visible = False
          Me.Items(AccessControlMenuItems.acmiNone).Visible = True
          Me.Items(AccessControlMenuItems.acmiReadOnly).Visible = True
          Me.Items(AccessControlMenuItems.acmiSupervisor).Visible = True
          Me.Items(AccessControlMenuItems.acmiUser).Visible = True
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class
