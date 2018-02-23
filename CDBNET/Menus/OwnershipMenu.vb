Public Class OwnershipMenu
  Inherits ContextMenuStrip

  Private mvParent As MaintenanceParentForm
  Private mvNodeInfo As OwnershipNodeInfo
  Public Event MenuActionCompleted(ByVal pItem As OwnershipMenuItems)

  Public Enum OwnershipMenuItems
    omiNewGroup
    omiAddUser
    omiAddDepartment
    omiChangeDefault
    omiChangeAccess
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New()
    MyBase.New()
    With mvMenuItems
      .Add(OwnershipMenuItems.omiNewGroup.ToString, New MenuToolbarCommand("NewGroup", ControlText.MnuNewGroup, OwnershipMenuItems.omiNewGroup))
      .Add(OwnershipMenuItems.omiAddUser.ToString, New MenuToolbarCommand("AddUser", ControlText.MnuAddUser, OwnershipMenuItems.omiAddUser))
      .Add(OwnershipMenuItems.omiAddDepartment.ToString, New MenuToolbarCommand("AddDepartment", ControlText.MnuAddDepartment, OwnershipMenuItems.omiAddDepartment))
      .Add(OwnershipMenuItems.omiChangeDefault.ToString, New MenuToolbarCommand("ChangeDefault", ControlText.MnuChangeDefault, OwnershipMenuItems.omiChangeDefault))
      .Add(OwnershipMenuItems.omiChangeAccess.ToString, New MenuToolbarCommand("ChangeAccess", ControlText.MnuChangeAccess, OwnershipMenuItems.omiChangeAccess))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
  End Sub

  Public Property OwnershipNodeInfo() As OwnershipNodeInfo
    Get
      Return mvNodeInfo
    End Get
    Set(ByVal Value As OwnershipNodeInfo)
      mvNodeInfo = Value
    End Set
  End Property

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As OwnershipMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, OwnershipMenuItems)
      Dim vCurrentTableName As String
      Dim vParams As New ParameterList(True)
      Dim vCriteriaList As New ParameterList
      Dim vForm As frmTableEntry = Nothing
      Dim vFormAppParams As frmApplicationParameters = Nothing

      Select Case vMenuItem
        Case OwnershipMenuItems.omiNewGroup
          vCurrentTableName = "ownership_groups"
          vParams("MaintenanceTableName") = vCurrentTableName
          vForm = New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vCurrentTableName, vParams, Nothing)
          vForm.Text = ControlText.TxtOwnGroup
        Case OwnershipMenuItems.omiAddUser
          vCurrentTableName = "ownership_group_users"
          vParams("MaintenanceTableName") = vCurrentTableName
          vParams("ValidFrom") = Today.ToShortDateString
          vParams("OwnershipGroup") = mvNodeInfo.OwnershipGroup
          vForm = New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vCurrentTableName, vParams, vCriteriaList)
          vForm.Text = ControlText.TxtOwnGroupUser
        Case OwnershipMenuItems.omiAddDepartment
          vCurrentTableName = "department_ownership_defaults"
          vParams("MaintenanceTableName") = vCurrentTableName
          vParams("OwnershipGroup") = mvNodeInfo.OwnershipGroup
          vForm = New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vCurrentTableName, vParams, vCriteriaList)
          vForm.Text = ControlText.TxtDepOwnDefault
        Case OwnershipMenuItems.omiChangeDefault
          vCurrentTableName = "department_ownership_defaults"
          vParams("MaintenanceTableName") = vCurrentTableName
          vParams("OwnershipGroup") = mvNodeInfo.OwnershipGroup
          vParams("Department") = mvNodeInfo.Department
          vParams("OwnershipAccessLevel") = mvNodeInfo.AcceesLevel
          vForm = New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, vCurrentTableName, vParams, vCriteriaList)
          vForm.Text = ControlText.TxtDepOwnDefault
        Case OwnershipMenuItems.omiChangeAccess
          vCurrentTableName = "ownership_group_users"
          vCriteriaList("Logname") = mvNodeInfo.User
          vParams("MaintenanceTableName") = vCurrentTableName
          vParams("ValidFrom") = Today.ToShortDateString
          vParams("OwnershipGroup") = mvNodeInfo.OwnershipGroup
          vParams("User") = mvNodeInfo.User
          vFormAppParams = New frmApplicationParameters(EditPanelInfo.OtherPanelTypes.optOwnershipMaintenance, Nothing, vParams, ControlText.TxtOwnGroupUser)
          If vFormAppParams.ShowDialog() = DialogResult.OK Then
            RaiseEvent MenuActionCompleted(vMenuItem)
          End If
      End Select
      If Not vMenuItem = OwnershipMenuItems.omiChangeAccess Then
        If vForm.ShowDialog() = DialogResult.OK Then
          RaiseEvent MenuActionCompleted(vMenuItem)
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub OwnershipMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    Dim vCursor As New BusyCursor
    Try
      Select Case mvNodeInfo.NodeType
        Case OwnershipNodeType.OwnershipGroupType
          Me.Items(OwnershipMenuItems.omiNewGroup).Visible = True
          Me.Items(OwnershipMenuItems.omiAddUser).Visible = False
          Me.Items(OwnershipMenuItems.omiAddDepartment).Visible = False
          Me.Items(OwnershipMenuItems.omiChangeAccess).Visible = False
          Me.Items(OwnershipMenuItems.omiChangeDefault).Visible = False
        Case OwnershipNodeType.OwnershipGroup, OwnershipNodeType.UserType, OwnershipNodeType.DepartmentType
          Me.Items(OwnershipMenuItems.omiNewGroup).Visible = True
          Me.Items(OwnershipMenuItems.omiAddUser).Visible = True
          Me.Items(OwnershipMenuItems.omiAddDepartment).Visible = True
          Me.Items(OwnershipMenuItems.omiChangeAccess).Visible = False
          Me.Items(OwnershipMenuItems.omiChangeDefault).Visible = False
        Case OwnershipNodeType.User
          Me.Items(OwnershipMenuItems.omiNewGroup).Visible = True
          Me.Items(OwnershipMenuItems.omiAddUser).Visible = False
          Me.Items(OwnershipMenuItems.omiAddDepartment).Visible = False
          Me.Items(OwnershipMenuItems.omiChangeAccess).Visible = True
          Me.Items(OwnershipMenuItems.omiChangeDefault).Visible = False
        Case OwnershipNodeType.Department
          Me.Items(OwnershipMenuItems.omiNewGroup).Visible = True
          Me.Items(OwnershipMenuItems.omiAddUser).Visible = False
          Me.Items(OwnershipMenuItems.omiAddDepartment).Visible = False
          Me.Items(OwnershipMenuItems.omiChangeAccess).Visible = False
          Me.Items(OwnershipMenuItems.omiChangeDefault).Visible = True
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class
